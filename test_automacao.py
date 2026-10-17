import csv
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from openpyxl import Workbook
from selenium.common.exceptions import TimeoutException

from automacao import (
    LEGACY_REPORT_FIELDS,
    REPORT_FIELDS,
    load_note_metadata,
    load_notes,
    normalize_note,
    write_note_descriptions,
    write_report,
)
from salesforce import (
    ResultadoNota,
    _abrir_compromisso_servico,
    _abrir_aba_relacionados,
    _baixar_arquivos_relacionados,
    _compromissos_relacionados,
    _extract_downloaded_images,
    _matching_work_order_result,
    _modal_image_rows,
    _modal_download_button,
    _processar_nota,
    _visible_search_box,
    _wait_for_stable_search_box,
    _buscar_ordem,
    consultar_notas,
    criar_navegador,
)


class NormalizeNoteTests(unittest.TestCase):
    def test_normalizes_integer_and_integral_float(self):
        self.assertEqual(normalize_note(12345, 2), "12345")
        self.assertEqual(normalize_note(12345.0, 2), "12345")

    def test_trims_text_and_ignores_blank_values(self):
        self.assertEqual(normalize_note(" 00123 ", 2), "00123")
        self.assertIsNone(normalize_note(" ", 2))
        self.assertIsNone(normalize_note(None, 2))

    def test_rejects_invalid_numbers(self):
        with self.assertRaises(ValueError):
            normalize_note(12.5, 2)
        with self.assertRaises(ValueError):
            normalize_note(True, 2)
        with self.assertRaisesRegex(ValueError, "somente algarismos"):
            normalize_note("../fora", 2)


class WorkbookTests(unittest.TestCase):
    def test_loads_column_case_insensitively_and_deduplicates(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "notas.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Notas sem evidências"
            sheet.append(["NOTA", "OBS"])
            sheet.append([123, "a"])
            sheet.append([123.0, "duplicada"])
            sheet.append([" 00456 ", "b"])
            sheet.append([None, "vazia"])
            workbook.save(path)

            self.assertEqual(
                load_notes(path, "Notas sem evidências", "Nota"),
                ["123", "00456"],
            )

    def test_reports_missing_sheet_or_column(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "notas.xlsx"
            workbook = Workbook()
            workbook.active.append(["Outra coluna"])
            workbook.save(path)

            with self.assertRaisesRegex(ValueError, "Aba"):
                load_notes(path, "Inexistente", "Nota")
            with self.assertRaisesRegex(ValueError, "Coluna"):
                load_notes(path, "Sheet", "Nota")

    def test_loads_metadata_and_writes_one_description_per_note(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "notas.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Notas sem evidências"
            sheet.append(
                [
                    "Nota",
                    "Texto code ação",
                    "Código de ação",
                    "Data encermto.",
                    "CenTrab respon.",
                    "Texto code ação",
                    "UTD",
                ]
            )
            sheet.append(
                [
                    123,
                    "Não Localizado",
                    "NE21",
                    datetime(2026, 9, 28),
                    "4DW501",
                    "Não Localizado",
                    "UTD-01",
                ]
            )
            workbook.save(path)

            metadata = load_note_metadata(
                path,
                "Notas sem evidências",
                "Nota",
                ["123"],
            )
            output = root / "Respostas"
            (output / "123").mkdir(parents=True)
            write_note_descriptions(output, ["123"], metadata)

            self.assertEqual(metadata["123"]["data_encerramento"], "28/09/2026")
            self.assertEqual(metadata["123"]["UTD"], "UTD-01")
            self.assertEqual(
                (output / "123" / "descricao.txt").read_text(encoding="utf-8"),
                "Nota: 123\n"
                "Texto code ação: Não Localizado\n"
                "Código de ação: NE21\n"
                "Data de encerramento: 28/09/2026\n"
                "Centro de trabalho responsável: 4DW501\n",
            )

    def test_rejects_conflicting_duplicate_description_columns(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "notas.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Notas sem evidências"
            sheet.append(
                [
                    "Nota",
                    "Texto code ação",
                    "Código de ação",
                    "Data encermto.",
                    "CenTrab respon.",
                    "Texto code ação",
                    "UTD",
                ]
            )
            sheet.append([123, "A", "NE01", None, "4DW501", "B", "UTD-01"])
            workbook.save(path)

            with self.assertRaisesRegex(ValueError, "duplicadas"):
                load_note_metadata(path, "Notas sem evidências", "Nota", ["123"])


class ReportTests(unittest.TestCase):
    def test_preserves_existing_results_for_notes_in_new_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "relatorio.csv"
            write_report(path, ["123", "456"])

            with path.open("r", encoding="utf-8-sig", newline="") as report_file:
                rows = list(csv.DictReader(report_file, delimiter=";"))
            rows[0]["status"] = "Concluída"
            rows[0]["motivo_nao_conclusao"] = "Acesso impedido"
            with path.open("w", encoding="utf-8-sig", newline="") as report_file:
                writer = csv.DictWriter(
                    report_file, fieldnames=rows[0].keys(), delimiter=";"
                )
                writer.writeheader()
                writer.writerows(rows)

            write_report(path, ["123", "789"])
            with path.open("r", encoding="utf-8-sig", newline="") as report_file:
                updated = list(csv.DictReader(report_file, delimiter=";"))

            self.assertEqual([row["nota"] for row in updated], ["123", "789"])
            self.assertEqual(updated[0]["status"], "Concluída")
            self.assertEqual(updated[0]["motivo_nao_conclusao"], "Acesso impedido")
            self.assertEqual(updated[1]["status"], "Pendente")

    def test_applies_salesforce_results_without_losing_other_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "relatorio.csv"
            write_report(path, ["123", "456"])
            update = {
                field: ""
                for field in REPORT_FIELDS
            }
            update.update(
                {
                    "nota": "123",
                    "UTD": "UTD-01",
                    "status": "Concluída",
                    "motivo_nao_conclusao": "Sem acesso",
                    "arquivos_baixados": "123_foto.jpg",
                }
            )
            write_report(path, ["123", "456"], {"123": update})

            with path.open("r", encoding="utf-8-sig", newline="") as report_file:
                rows = list(csv.DictReader(report_file, delimiter=";"))

            self.assertEqual(rows[0]["status"], "Concluída")
            self.assertEqual(rows[0]["UTD"], "UTD-01")
            self.assertEqual(rows[0]["motivo_nao_conclusao"], "Sem acesso")
            self.assertEqual(rows[0]["arquivos_baixados"], "123_foto.jpg")
            self.assertEqual(rows[1]["status"], "Pendente")

    def test_migrates_legacy_report_and_preserves_existing_status(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "relatorio.csv"
            with path.open("w", encoding="utf-8-sig", newline="") as report_file:
                writer = csv.DictWriter(
                    report_file,
                    fieldnames=LEGACY_REPORT_FIELDS,
                    delimiter=";",
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "nota": "123",
                        "status": "Concluída",
                        "motivo_nao_conclusao": "Sem acesso",
                        "arquivos_baixados": "123_foto.jpg",
                        "erro": "",
                    }
                )

            write_report(
                path,
                ["123"],
                {"123": {"data_encerramento": "28/09/2026"}},
            )

            with path.open("r", encoding="utf-8-sig", newline="") as report_file:
                rows = list(csv.DictReader(report_file, delimiter=";"))

            self.assertEqual(rows[0]["status"], "Concluída")
            self.assertEqual(rows[0]["motivo_nao_conclusao"], "Sem acesso")
            self.assertEqual(rows[0]["data_encerramento"], "28/09/2026")
            self.assertEqual(rows[0]["centro_trabalho"], "")


class SalesforceLifecycleTests(unittest.TestCase):
    def test_keeps_chrome_open_when_login_does_not_finish(self):
        driver = Mock()
        with (
            patch("salesforce.criar_navegador", return_value=driver),
            patch(
                "salesforce._wait_for_login",
                side_effect=RuntimeError("login incompleto"),
            ),
        ):
            with self.assertRaisesRegex(RuntimeError, "login incompleto"):
                consultar_notas([], "https://example.my.salesforce.com/", Path("."), None)

        driver.service.stop.assert_called_once()
        driver.quit.assert_not_called()

    def test_closes_chrome_after_successful_authenticated_run(self):
        driver = Mock()
        with (
            patch("salesforce.criar_navegador", return_value=driver),
            patch("salesforce._wait_for_login"),
            patch(
                "salesforce._processar_nota",
                return_value=ResultadoNota(nota="123", status="Concluída"),
            ),
        ):
            results = consultar_notas(
                ["123"],
                "https://example.my.salesforce.com/",
                Path("."),
                None,
            )

        self.assertEqual(results[0].status, "Concluída")
        driver.quit.assert_called_once()
        driver.service.stop.assert_not_called()

    def test_chrome_uses_a_detached_profile_for_manual_login(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            driver = Mock()
            with (
                patch.dict("os.environ", {"LOCALAPPDATA": str(root)}),
                patch("salesforce.webdriver.Chrome", return_value=driver) as chrome,
            ):
                criar_navegador(root / "downloads", root / "chrome-profile")

            options = chrome.call_args.kwargs["options"]
            self.assertTrue(options.experimental_options["detach"])


class SalesforceResultTests(unittest.TestCase):
    def test_search_result_is_activated_with_javascript_click(self):
        search_button = Mock()
        search_box = Mock()
        search_result = Mock()
        first_wait = Mock()
        first_wait.until.return_value = search_button
        result_wait = Mock()
        result_wait.until.return_value = search_result
        navigation_wait = Mock()
        navigation_wait.until.side_effect = lambda condition: condition(driver)
        driver = Mock()
        driver.current_url = "https://example.lightning.force.com/lightning/r/0WOR123/view"
        driver.find_elements.return_value = [
            SimpleNamespace(text="Ordem de Serviço COE_CS1030177305")
        ]

        with (
            patch(
                "salesforce.WebDriverWait",
                side_effect=[result_wait, navigation_wait],
            ),
            patch(
                "salesforce._wait_for_stable_search_box",
                side_effect=[False, False, search_box],
            ),
        ):
            _buscar_ordem(driver, first_wait, "1030177305")

        search_button.click.assert_called_once()
        search_box.click.assert_not_called()
        search_box.send_keys.assert_called_once_with("1030177305")
        self.assertEqual(driver.execute_script.call_count, 2)
        self.assertIs(driver.execute_script.call_args_list[0].args[1], search_button)
        self.assertIs(driver.execute_script.call_args_list[1].args[1], search_result)
        search_result.click.assert_not_called()

    def test_opens_related_tab_and_waits_for_it_to_be_selected(self):
        selected = {"value": False}
        tab = Mock()
        tab.text = "Relacionado"
        tab.get_attribute.side_effect = lambda name: (
            "true" if selected["value"] else "false"
        ) if name == "aria-selected" else None
        download_button = Mock()
        download_button.is_displayed.return_value = True
        download_button.is_enabled.return_value = True
        driver = Mock()
        driver.find_elements.side_effect = [[tab], [tab]]

        def click_related_tab(_script, clicked_tab):
            self.assertIs(clicked_tab, tab)
            selected["value"] = True

        driver.execute_script.side_effect = click_related_tab
        wait = Mock()
        wait.until.side_effect = lambda condition: condition(driver)

        _abrir_aba_relacionados(driver, wait, "1030177305")

        self.assertTrue(selected["value"])
        self.assertEqual(driver.find_elements.call_count, 2)

    def test_finds_service_appointments_in_the_work_order_related_card(self):
        appointment_url = (
            "https://example.lightning.force.com/"
            "lightning/r/08pRo000000000001/view"
        )
        appointment_link = Mock()
        appointment_link.get_attribute.return_value = appointment_url
        card = Mock()
        card.find_elements.side_effect = lambda _by, selector: (
            [appointment_link]
            if selector == ".//a[contains(@href, '/lightning/r/08p')]"
            else []
        )
        driver = Mock()
        driver.find_elements.return_value = [card]
        wait = Mock()
        wait.until.side_effect = lambda condition: condition(driver)

        result = _compromissos_relacionados(wait, "1030177305")

        self.assertEqual(result, [appointment_url])

    def test_returns_empty_when_related_work_order_has_no_appointments(self):
        heading = SimpleNamespace(text="Compromissos de serviços (0)")
        card = Mock()
        card.find_elements.side_effect = lambda _by, selector: (
            [heading]
            if selector == ".//h2[contains(., 'Compromissos de serviços')]"
            else []
        )
        driver = Mock()
        driver.find_elements.return_value = [card]
        wait = Mock()
        wait.until.side_effect = lambda condition: condition(driver)

        self.assertEqual(
            _compromissos_relacionados(wait, "1030177305"),
            [],
        )

    def test_opens_service_appointment_from_the_related_card(self):
        appointment_id = "08pRo000000000001"
        appointment_url = (
            f"https://example.lightning.force.com/lightning/r/{appointment_id}/view"
        )
        appointment_link = Mock()
        appointment_link.get_attribute.return_value = appointment_url
        appointment_link.is_displayed.return_value = True
        card = Mock()
        card.find_elements.return_value = [appointment_link]
        driver = Mock()
        driver.current_url = (
            "https://example.lightning.force.com/lightning/r/WorkOrder/123/view"
        )
        driver.find_elements.return_value = [card]

        def click_link(_script, clicked_link):
            self.assertIs(clicked_link, appointment_link)
            driver.current_url = (
                "https://example.lightning.force.com/lightning/r/"
                f"ServiceAppointment/{appointment_id}/view"
            )

        driver.execute_script.side_effect = click_link
        wait = Mock()
        wait.until.side_effect = lambda condition: condition(driver)

        _abrir_compromisso_servico(
            driver,
            wait,
            "1030177305",
            appointment_url,
        )

        self.assertIn(appointment_id, driver.current_url)
        driver.execute_script.assert_called_once()

    def test_search_box_searches_open_shadow_roots(self):
        visible_input = Mock()
        driver = Mock()
        driver.execute_script.return_value = visible_input

        self.assertIs(_visible_search_box(driver), visible_input)
        script = driver.execute_script.call_args.args[0]
        self.assertIn("element.shadowRoot", script)
        self.assertIn("type === 'search'", script)
        self.assertIn("rect.width > 0", script)
        self.assertIn("document.elementFromPoint", script)
        self.assertIn("topElement.shadowRoot.elementFromPoint", script)
        self.assertIn("return topElement === element", script)
        self.assertIn("isAccessibleFromPage(element)", script)

    def test_search_box_waits_for_stable_geometry(self):
        search_box = Mock()
        driver = Mock()
        driver.execute_script.side_effect = [
            [0, 0, 0, 0],
            [321, 52, 664, 32],
            [321, 52, 664, 32],
            [321, 52, 664, 32],
        ]

        def until_stable(condition):
            for _ in range(4):
                result = condition(driver)
                if result:
                    return result
            raise TimeoutException()

        with (
            patch("salesforce.WebDriverWait") as wait_class,
            patch(
                "salesforce._visible_search_box",
                return_value=search_box,
            ) as visible_search_box,
        ):
            wait_class.return_value.until.side_effect = until_stable

            result = _wait_for_stable_search_box(driver, timeout=2)

        self.assertIs(result, search_box)
        self.assertEqual(visible_search_box.call_count, 4)

    def test_search_box_does_not_select_input_outside_dialog(self):
        driver = Mock()
        driver.execute_script.return_value = None

        self.assertFalse(_visible_search_box(driver))

    def test_search_box_excludes_salesforce_object_type_picker(self):
        driver = Mock()
        driver.execute_script.return_value = None

        self.assertFalse(_visible_search_box(driver))
        self.assertIn(
            "by object type",
            driver.execute_script.call_args.args[0],
        )

    def test_ignores_hidden_search_results(self):
        hidden = SimpleNamespace(
            is_displayed=lambda: False,
            is_enabled=lambda: True,
            text="04596443 Ordem de trabalho • COE_CS1030177305",
        )
        visible = SimpleNamespace(
            is_displayed=lambda: True,
            is_enabled=lambda: True,
            text="04596443 Ordem de trabalho • COE_CS1030177305",
        )
        driver = Mock()
        driver.find_elements.return_value = [hidden, visible]

        self.assertIs(
            _matching_work_order_result(driver, "1030177305"),
            visible,
        )

    def test_extracts_all_images_from_zip_and_ignores_other_files(self):
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory)
            archive_path = destination / "download.zip"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr("equipe/foto1.JPG", b"image one")
                archive.writestr("equipe/foto2.png", b"image two")
                archive.writestr("equipe/relatorio.pdf", b"document")

            names = _extract_downloaded_images(
                [archive_path],
                "1030177305",
                destination,
            )

            self.assertEqual(
                names,
                ["1030177305_foto1.JPG", "1030177305_foto2.png"],
            )
            self.assertEqual(
                (destination / names[0]).read_bytes(),
                b"image one",
            )
            self.assertEqual(
                (destination / names[1]).read_bytes(),
                b"image two",
            )
            self.assertFalse(archive_path.exists())

    def test_empty_zip_returns_no_images_without_failing(self):
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory)
            archive_path = destination / "download.zip"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr("readme.txt", "no photos")

            self.assertEqual(
                _extract_downloaded_images(
                    [archive_path],
                    "1030177305",
                    destination,
                ),
                [],
            )
            self.assertFalse(archive_path.exists())

    def test_corrupt_zip_is_reported_instead_of_marked_as_no_images(self):
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory)
            archive_path = destination / "download.zip"
            archive_path.write_bytes(b"not a zip archive")

            with self.assertRaisesRegex(RuntimeError, "Não foi possível extrair"):
                _extract_downloaded_images(
                    [archive_path],
                    "1030177305",
                    destination,
                )

            self.assertTrue(archive_path.exists())

    def test_empty_attachments_are_reported_as_no_images(self):
        with tempfile.TemporaryDirectory() as directory:
            driver = Mock()
            driver.current_url = (
                "https://example.lightning.force.com/lightning/r/WorkOrder/123/view"
            )
            with (
                patch("salesforce.WebDriverWait", return_value=Mock()),
                patch("salesforce._buscar_ordem"),
                patch("salesforce._abrir_aba_relacionados"),
                patch(
                    "salesforce._compromissos_relacionados",
                    return_value=[
                        "https://example.lightning.force.com/lightning/r/"
                        "08pRo000000000001/view"
                    ],
                ),
                patch("salesforce._abrir_compromisso_servico"),
                patch("salesforce._ler_motivo", return_value=""),
                patch("salesforce._baixar_arquivos_relacionados", return_value=[]),
            ):
                result = _processar_nota(
                    driver,
                    Path(directory),
                    "1030177305",
                )

            self.assertEqual(result.status, "Sem imagens")
            self.assertEqual(result.arquivos, [])
            self.assertEqual(result.erro, "")

    def test_opens_related_service_appointment_before_downloading_files(self):
        with tempfile.TemporaryDirectory() as directory:
            driver = Mock()
            driver.current_url = (
                "https://example.lightning.force.com/lightning/r/WorkOrder/123/view"
            )
            events = []
            with (
                patch("salesforce.WebDriverWait", return_value=Mock()),
                patch("salesforce._buscar_ordem"),
                patch(
                    "salesforce._abrir_aba_relacionados",
                    side_effect=lambda *_args: events.append("related"),
                ) as open_related,
                patch(
                    "salesforce._compromissos_relacionados",
                    return_value=[
                        "https://example.lightning.force.com/lightning/r/"
                        "08pRo000000000001/view"
                    ],
                ),
                patch(
                    "salesforce._abrir_compromisso_servico",
                    side_effect=lambda *_args: events.append("appointment"),
                ) as open_appointment,
                patch("salesforce._ler_motivo", return_value=""),
                patch(
                    "salesforce._baixar_arquivos_relacionados",
                    side_effect=lambda *_args: events.append("download")
                    or ["1030177305_foto.jpg"],
                ) as download_files,
            ):
                result = _processar_nota(
                    driver,
                    Path(directory),
                    "1030177305",
                )

            self.assertEqual(result.status, "Concluída")
            self.assertEqual(result.arquivos, ["1030177305_foto.jpg"])
            self.assertEqual(
                events,
                ["related", "appointment", "download"],
            )
            open_related.assert_called_once()
            open_appointment.assert_called_once()
            download_files.assert_called_once()
            driver.get.assert_not_called()

    def test_empty_file_modal_returns_no_images_without_clicking_download(self):
        toolbar_button = Mock()
        close_button = Mock()
        modal = Mock()
        modal.is_displayed.return_value = True
        modal.text = "Nenhum arquivo"

        def modal_elements(_by, selector):
            if selector == ".//button[contains(., 'Cancelar e fechar')]":
                return [close_button]
            return []

        modal.find_elements.side_effect = modal_elements

        initial_wait = Mock()
        initial_wait.until.return_value = toolbar_button
        modal_wait = Mock()
        modal_wait.until.side_effect = lambda condition: condition(driver)
        file_list_wait = Mock()
        file_list_wait.until.side_effect = lambda condition: condition(driver)
        driver = Mock()
        driver.find_elements.return_value = [modal]
        driver.execute_script.return_value = []

        with (
            tempfile.TemporaryDirectory() as directory,
            patch(
                "salesforce.WebDriverWait",
                side_effect=[modal_wait, file_list_wait],
            ),
        ):
            result = _baixar_arquivos_relacionados(
                driver,
                initial_wait,
                "1030177305",
                Path(directory),
            )
        self.assertEqual(result, [])
        toolbar_button.click.assert_called_once()
        driver.execute_script.assert_any_call(
            "arguments[0].click()",
            close_button,
        )
        self.assertEqual(driver.execute_script.call_count, 3)

    def test_modal_image_rows_are_found_inside_shadow_dom(self):
        checkbox = Mock()
        modal = Mock()
        driver = Mock()
        driver.execute_script.return_value = [
            {"checkbox": checkbox, "filename": "10001129.jpg"},
            {"checkbox": Mock(), "filename": "10001130.jpg"},
        ]

        rows = _modal_image_rows(driver, modal)

        self.assertEqual(
            rows,
            [(checkbox, "10001129.jpg"), (driver.execute_script.return_value[1]["checkbox"], "10001130.jpg")],
        )
        script = driver.execute_script.call_args.args[0]
        self.assertIn("element.shadowRoot", script)
        self.assertIn("input[type=\"checkbox\"]", script)

    def test_modal_download_button_is_found_inside_shadow_dom(self):
        modal = Mock()
        button = Mock()
        driver = Mock()
        driver.execute_script.return_value = button

        result = _modal_download_button(driver, modal)

        self.assertIs(result, button)
        script = driver.execute_script.call_args.args[0]
        self.assertIn("element.shadowRoot", script)
        self.assertIn("label === 'baixar'", script)

    def test_waits_for_modal_file_rows_before_selecting_and_downloading(self):
        heading = Mock()
        heading.is_displayed.return_value = True
        toolbar_button = Mock()
        checkbox = Mock()
        checkbox.is_selected.return_value = False
        photo_rows = [{"checkbox": checkbox, "filename": "equipe/foto.jpg"}]
        modal = Mock()
        modal.is_displayed.return_value = True
        modal.text = "Arquivo Tipo Baixar"
        modal.find_elements.side_effect = lambda _by, selector: (
            [download_button]
            if selector == ".//button[normalize-space(.)='Baixar']"
            else []
        )
        download_button = Mock()

        def create_zip():
            with zipfile.ZipFile(Path(directory) / "download.zip", "w") as archive:
                archive.writestr("equipe/foto.jpg", b"image")
                archive.writestr("equipe/relatorio.pdf", b"document")

        initial_wait = Mock()
        initial_wait.until.return_value = toolbar_button
        modal_wait = Mock()
        modal_wait.until.side_effect = lambda condition: condition(driver)
        file_list_wait = Mock()
        polls = {"count": 0}

        def wait_for_rows(condition):
            result = condition(driver)
            if not result:
                polls["count"] += 1
                result = condition(driver)
            return result

        file_list_wait.until.side_effect = wait_for_rows
        download_wait = Mock()
        download_wait.until.return_value = download_button
        driver = Mock()
        driver.find_elements.return_value = [modal]

        row_scans = {"count": 0}

        def execute_script(script, target):
            if "const matches = []" in script:
                row_scans["count"] += 1
                return [] if row_scans["count"] == 1 else photo_rows
            if target is download_button:
                create_zip()

        driver.execute_script.side_effect = execute_script

        with (
            tempfile.TemporaryDirectory() as directory,
            patch(
                "salesforce.WebDriverWait",
                side_effect=[modal_wait, file_list_wait, download_wait],
            ),
            patch("salesforce.time.monotonic", side_effect=[0, 0, 0, 3, 3]),
            patch("salesforce.time.sleep"),
        ):
            result = _baixar_arquivos_relacionados(
                driver,
                initial_wait,
                "1030177305",
                Path(directory),
            )

            self.assertEqual(result, ["1030177305_foto.jpg"])
            self.assertEqual(
                (Path(directory) / "1030177305_foto.jpg").read_bytes(),
                b"image",
            )
            self.assertFalse((Path(directory) / "download.zip").exists())

        driver.execute_script.assert_any_call(
            "arguments[0].click()",
            checkbox,
        )
        driver.execute_script.assert_any_call(
            "arguments[0].click()",
            download_button,
        )
        self.assertEqual(polls["count"], 1)
        self.assertGreaterEqual(row_scans["count"], 3)

    def test_direct_download_from_files_button_is_extracted_without_modal(self):
        toolbar_button = Mock()

        def create_zip():
            with zipfile.ZipFile(Path(directory) / "download.zip", "w") as archive:
                archive.writestr("equipe/foto.jpg", b"image")
                archive.writestr("equipe/relatorio.pdf", b"document")

        toolbar_button.click.side_effect = create_zip
        initial_wait = Mock()
        initial_wait.until.return_value = toolbar_button
        download_wait = Mock()
        download_wait.until.side_effect = lambda condition: condition(driver)
        driver = Mock()
        driver.find_elements.return_value = []

        with (
            tempfile.TemporaryDirectory() as directory,
            patch("salesforce.WebDriverWait", return_value=download_wait),
            patch("salesforce.time.monotonic", side_effect=[0, 0, 0, 3, 3]),
            patch("salesforce.time.sleep"),
        ):
            result = _baixar_arquivos_relacionados(
                driver,
                initial_wait,
                "1030177305",
                Path(directory),
            )

            self.assertEqual(result, ["1030177305_foto.jpg"])
            self.assertEqual(
                (Path(directory) / "1030177305_foto.jpg").read_bytes(),
                b"image",
            )
            self.assertFalse((Path(directory) / "download.zip").exists())

        toolbar_button.click.assert_called_once()
        driver.find_elements.assert_called_once()

    def test_empty_file_table_with_headers_returns_no_images(self):
        toolbar_button = Mock()
        close_button = Mock()
        modal = Mock()
        modal.is_displayed.return_value = True
        modal.text = "Arquivo Tipo Nenhum arquivo Baixar"
        modal.find_elements.side_effect = lambda _by, selector: (
            [Mock()]
            if selector == ".//button[normalize-space(.)='Baixar']"
            else [close_button]
            if selector == ".//button[contains(., 'Cancelar e fechar')]"
            else []
        )
        initial_wait = Mock()
        initial_wait.until.return_value = toolbar_button
        modal_wait = Mock()
        modal_wait.until.side_effect = lambda condition: condition(driver)
        file_list_wait = Mock()
        file_list_wait.until.side_effect = lambda condition: condition(driver)
        driver = Mock()
        driver.find_elements.return_value = [modal]
        driver.execute_script.return_value = []

        with (
            tempfile.TemporaryDirectory() as directory,
            patch(
                "salesforce.WebDriverWait",
                side_effect=[modal_wait, file_list_wait],
            ),
        ):
            result = _baixar_arquivos_relacionados(
                driver,
                initial_wait,
                "1030177305",
                Path(directory),
            )

        self.assertEqual(result, [])
        self.assertEqual(driver.execute_script.call_count, 3)
        driver.execute_script.assert_any_call(
            "arguments[0].click()",
            close_button,
        )

    def test_reports_failure_to_open_the_related_tab(self):
        with tempfile.TemporaryDirectory() as directory:
            driver = Mock()
            with (
                patch("salesforce.WebDriverWait", return_value=Mock()),
                patch("salesforce._buscar_ordem"),
                patch(
                    "salesforce._abrir_aba_relacionados",
                    side_effect=RuntimeError("aba Relacionados indisponível"),
                ),
                patch("salesforce._compromissos_relacionados"),
                patch("salesforce._ler_motivo", return_value=""),
                patch(
                    "salesforce._baixar_arquivos_relacionados"
                ) as download_files,
            ):
                result = _processar_nota(
                    driver,
                    Path(directory),
                    "1030177305",
                )

            self.assertEqual(result.status, "Erro")
            self.assertEqual(result.erro, "aba Relacionados indisponível")
            download_files.assert_not_called()


if __name__ == "__main__":
    unittest.main()
