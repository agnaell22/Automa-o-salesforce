import hashlib
import json
import os
import tempfile
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait


LOGIN_TIMEOUT_SECONDS = 300
PAGE_TIMEOUT_SECONDS = 30
CAPTION_FIELD_TIMEOUT_SECONDS = 3
SUCCESSFUL_SALESFORCE_STATUSES = frozenset({"Concluída"})
IMAGE_EXTENSIONS = frozenset(
    {".jpg", ".jpeg", ".png", ".gif", ".webp", ".heic", ".bmp"}
)


def create_responses_archive(output_dir: Path) -> Path:
    if not output_dir.is_dir():
        raise FileNotFoundError(f"Pasta de respostas não encontrada: {output_dir}")

    archive_path = output_dir.parent / f"{output_dir.name}.zip"
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{archive_path.stem}_",
        suffix=".tmp",
        dir=archive_path.parent,
    )
    os.close(file_descriptor)
    try:
        with zipfile.ZipFile(
            temporary_name,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as archive:
            files = sorted(
                path
                for path in output_dir.rglob("*")
                if path.is_file() and not path.is_symlink()
            )
            for path in files:
                relative_path = path.relative_to(output_dir).as_posix()
                info = zipfile.ZipInfo(
                    relative_path,
                    date_time=(1980, 1, 1, 0, 0, 0),
                )
                info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(info, path.read_bytes())
        os.replace(temporary_name, archive_path)
    finally:
        if os.path.exists(temporary_name):
            os.remove(temporary_name)
    return archive_path


def _state_path(output_dir: Path) -> Path:
    return output_dir.parent / f".{output_dir.name}_whatsapp_state.json"


def _load_state(state_path: Path) -> dict[str, Any]:
    if not state_path.exists():
        return {"notes": {}, "archive_sha256": ""}
    try:
        state = json.loads(state_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RuntimeError(
            f"Não foi possível ler o estado de envios do WhatsApp: {error}"
        ) from error
    if (
        not isinstance(state, dict)
        or not isinstance(state.get("notes"), dict)
        or not isinstance(state.get("archive_sha256", ""), str)
    ):
        raise RuntimeError(
            f"O estado de envios do WhatsApp está inválido: {state_path}"
        )
    return state


def _save_state(state_path: Path, state: dict[str, Any]) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{state_path.stem}_",
        suffix=".tmp",
        dir=state_path.parent,
    )
    try:
        with os.fdopen(file_descriptor, "w", encoding="utf-8", newline="\n") as file:
            json.dump(state, file, ensure_ascii=False, indent=2)
            file.write("\n")
        os.replace(temporary_name, state_path)
    finally:
        if os.path.exists(temporary_name):
            os.remove(temporary_name)


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _note_message(description_path: Path, report_row: dict[str, str]) -> str:
    if not description_path.is_file():
        raise FileNotFoundError(f"Descrição da nota não encontrada: {description_path}")
    message = description_path.read_text(encoding="utf-8").strip()
    reason = (report_row.get("motivo_nao_conclusao") or "").strip()
    if reason:
        message = f"{message}\nMotivo do impedimento: {reason}"
    return message


def _note_fingerprint(
    note: str,
    message: str,
    images: list[Path],
) -> str:
    digest = hashlib.sha256()
    digest.update(note.encode("utf-8"))
    digest.update(message.encode("utf-8"))
    for image in images:
        digest.update(image.name.encode("utf-8"))
        with image.open("rb") as file:
            for chunk in iter(lambda: file.read(1024 * 1024), b""):
                digest.update(chunk)
    return digest.hexdigest()


def _visible(element: WebElement) -> bool:
    try:
        return element.is_displayed()
    except WebDriverException:
        return False


def _search_button(driver: webdriver.Chrome) -> WebElement | bool:
    for button in driver.find_elements(By.CSS_SELECTOR, "button"):
        label = " ".join(
            (
                button.text or "",
                button.get_attribute("aria-label") or "",
                button.get_attribute("title") or "",
                button.get_attribute("data-testid") or "",
            )
        ).casefold()
        if _visible(button) and any(word in label for word in ("search", "pesquisar")):
            return button
    return False


def _search_input(driver: webdriver.Chrome) -> WebElement | bool:
    for element in driver.find_elements(
        By.CSS_SELECTOR,
        "input, [contenteditable='true']",
    ):
        label = " ".join(
            (
                element.get_attribute("aria-label") or "",
                element.get_attribute("placeholder") or "",
                element.get_attribute("data-testid") or "",
                element.get_attribute("title") or "",
            )
        ).casefold()
        if (
            _visible(element)
            and any(word in label for word in ("search", "pesquisar", "buscar"))
        ):
            return element
    return False


def _wait_for_search_input(driver: webdriver.Chrome) -> WebElement:
    def search_control(browser: webdriver.Chrome) -> tuple[str, WebElement] | bool:
        search_input = _search_input(browser)
        if search_input:
            return "input", search_input
        search_button = _search_button(browser)
        if search_button:
            return "button", search_button
        return False

    try:
        control = WebDriverWait(driver, LOGIN_TIMEOUT_SECONDS).until(search_control)
        if not control:
            raise TimeoutException("A busca do WhatsApp não apareceu.")
        control_type, control_element = control
        if control_type == "input":
            return control_element
        control_element.click()
        search_input = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            _search_input
        )
        if not search_input:
            raise TimeoutException("O campo de busca do WhatsApp não apareceu.")
        return search_input
    except TimeoutException as error:
        raise RuntimeError(
            "O WhatsApp Web não ficou pronto. Leia o QR Code e aguarde a lista "
            "de conversas carregar no Chrome."
        ) from error


def _xpath_literal(value: str) -> str:
    if "'" not in value:
        return f"'{value}'"
    if '"' not in value:
        return f'"{value}"'
    parts = value.split("'")
    return "concat(" + ', "\'", '.join(f"'{part}'" for part in parts) + ")"


def _open_group(driver: webdriver.Chrome, group_name: str) -> None:
    search_input = _wait_for_search_input(driver)
    try:
        search_input.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].focus()", search_input)
    search_input.send_keys(Keys.CONTROL, "a")
    search_input.send_keys(group_name)
    literal = _xpath_literal(group_name)

    def matching_result(browser: webdriver.Chrome) -> WebElement | bool:
        result_queries = (
            "//*[@data-testid='cell-frame-container' or @role='listitem']"
            f"[contains(@title, {literal}) or contains(@aria-label, {literal}) "
            f"or contains(normalize-space(.), {literal}) "
            f"or .//*[contains(@title, {literal}) "
            f"or contains(@aria-label, {literal}) "
            f"or contains(normalize-space(.), {literal})]]",
            "//*[contains(@title, "
            f"{literal}) or contains(@aria-label, {literal}) "
            f"or contains(normalize-space(.), {literal})]",
        )
        for query in result_queries:
            for element in browser.find_elements(By.XPATH, query):
                if _visible(element):
                    row = browser.execute_script(
                        """
                        return arguments[0].closest(
                            '[data-testid="cell-frame-container"], [role="listitem"]'
                        ) || arguments[0];
                        """,
                        element,
                    )
                    if row and _visible(row):
                        return row
        return False

    try:
        result = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(matching_result)
        if not result:
            raise TimeoutException("O grupo não apareceu na busca.")
    except TimeoutException as error:
        raise RuntimeError(
            f"O WhatsApp Web não exibiu um resultado visível para o grupo "
            f"'{group_name}'. Confirme que a conta conectada ao WhatsApp Web "
            "já participa desse grupo e que o nome corresponde ao cadastrado."
        ) from error

    def matching_header(browser: webdriver.Chrome) -> bool:
        for element in browser.find_elements(
            By.CSS_SELECTOR,
            "header [title], header [aria-label], header h1, header h2, "
            "header span, [data-testid='conversation-info-header'] "
            "[title], [data-testid='conversation-info-header'] "
            "[aria-label], [data-testid='conversation-info-header'] span",
        ):
            if not _visible(element):
                continue
            if any(
                element.get_attribute(attribute) == group_name
                for attribute in ("title", "aria-label")
            ):
                return True
            text_lines = {
                " ".join(line.split())
                for line in (element.text or "").splitlines()
                if line.strip()
            }
            if group_name in text_lines:
                return True
        return False

    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});",
            result,
        )
        try:
            result.click()
        except ElementClickInterceptedException:
            driver.execute_script(
                """
                arguments[0].click();
                """,
                result,
            )
        try:
            WebDriverWait(driver, 3).until(matching_header)
        except TimeoutException:
            driver.execute_script(
                "arguments[0].click();",
                result,
            )
            try:
                WebDriverWait(driver, 5).until(matching_header)
            except TimeoutException:
                diagnostics = driver.execute_script(
                    """
                    const visible = (element) =>
                        element.getClientRects().length > 0 &&
                        getComputedStyle(element).visibility !== 'hidden';
                    const describe = (element) => ({
                        testId: element.getAttribute('data-testid'),
                        role: element.getAttribute('role'),
                        title: element.getAttribute('title'),
                        aria: element.getAttribute('aria-label'),
                        text: (element.innerText || '').trim().slice(0, 160)
                    });
                    return JSON.stringify({
                        selectedRow: describe(arguments[0]),
                        headers: [...document.querySelectorAll(
                            'header, [data-testid="conversation-info-header"]'
                        )].filter(visible).slice(0, 4).map(describe),
                        visibleRows: [...document.querySelectorAll(
                            '[data-testid="cell-frame-container"], [role="listitem"]'
                        )].filter(visible).slice(0, 6).map(describe)
                    });
                    """,
                    result,
                )
                raise TimeoutException(str(diagnostics))
    except TimeoutException as error:
        raise RuntimeError(
            f"O WhatsApp Web exibiu o resultado de '{group_name}', mas não abriu "
            "a conversa nem confirmou o cabeçalho após clicar na linha do grupo. "
            f"Nenhuma mensagem foi enviada. Diagnóstico da interface: {error.msg}"
        ) from error


def _confirm(prompt: str, input_fn: Callable[[str], str]) -> bool:
    try:
        answer = input_fn(f"{prompt} [s/N]: ").strip().casefold()
    except EOFError as error:
        raise RuntimeError("A confirmação de envio foi interrompida.") from error
    return answer in {"s", "sim", "y", "yes"}


def _find_file_input(
    driver: webdriver.Chrome,
    *,
    images_only: bool,
) -> WebElement | bool:
    candidates: list[tuple[WebElement, str]] = []
    for element in driver.find_elements(By.CSS_SELECTOR, "input[type='file']"):
        accept = (element.get_attribute("accept") or "").casefold()
        is_image_input = "image" in accept or any(
            extension in accept for extension in IMAGE_EXTENSIONS
        )
        if is_image_input == images_only:
            candidates.append((element, accept))
    if images_only:
        for element, accept in candidates:
            if "video/" in accept:
                return element
    if candidates:
        return candidates[0][0]
    return False


def _media_dialog(driver: webdriver.Chrome) -> WebElement | bool:
    for dialog in driver.find_elements(By.XPATH, "//*[@role='dialog']"):
        if _visible(dialog):
            return dialog
    for preview in driver.find_elements(
        By.CSS_SELECTOR,
        "[data-testid*='media-editor'], [data-testid*='media-preview']",
    ):
        if not _visible(preview):
            continue
        if preview.get_attribute("data-testid") == "media-editor-canvas":
            editor = driver.execute_script(
                """
                const visible = (element) =>
                    element.getClientRects().length > 0 &&
                    getComputedStyle(element).visibility !== 'hidden';
                const label = (element) => [
                    element.innerText,
                    element.getAttribute('aria-label'),
                    element.getAttribute('title'),
                    element.getAttribute('data-testid'),
                    element.getAttribute('data-icon'),
                    ...[...element.querySelectorAll('[data-icon]')]
                        .map((icon) => icon.getAttribute('data-icon'))
                ].filter(Boolean).join(' ').toLowerCase();
                let node = arguments[0];
                for (let depth = 0; node && node !== document.body && depth < 10; depth++) {
                    const sendButton = [...node.querySelectorAll(
                        'button, [role="button"], [data-testid*="send"], [data-icon="send"]'
                    )].find((button) =>
                        visible(button) && /send|enviar/.test(label(button))
                    );
                    if (sendButton) return node;
                    node = node.parentElement;
                }
                return arguments[0].parentElement || arguments[0];
                """,
                preview,
            )
            if editor and _visible(editor):
                return editor
        elif _send_button(preview):
            return preview

    caption_labels = ("legenda", "caption")
    for element in driver.find_elements(
        By.CSS_SELECTOR,
        "textarea, [contenteditable='true'], input, [data-placeholder]",
    ):
        if not _visible(element):
            continue
        label = " ".join(
            (
                element.get_attribute("aria-label") or "",
                element.get_attribute("placeholder") or "",
                element.get_attribute("title") or "",
                element.get_attribute("data-placeholder") or "",
                element.get_attribute("data-testid") or "",
            )
        ).casefold()
        if not any(caption_label in label for caption_label in caption_labels):
            continue
        container = driver.execute_script(
            """
            let node = arguments[0].parentElement;
            while (node && node !== document.body) {
                const sendButton = [...node.querySelectorAll(
                    'button, [role="button"], [data-icon]'
                )].find((button) => {
                    const label = [
                        button.innerText,
                        button.getAttribute('aria-label'),
                        button.getAttribute('title'),
                        button.getAttribute('data-testid'),
                        button.getAttribute('data-icon')
                    ].filter(Boolean).join(' ').toLowerCase();
                    return /send|enviar/.test(label) &&
                        button.getClientRects().length > 0;
                });
                if (sendButton) return node;
                node = node.parentElement;
            }
            return arguments[0].parentElement;
            """,
            element,
        )
        if container and _visible(container):
            return container

    send_container = driver.execute_script(
        """
        const visible = (element) =>
            element.getClientRects().length > 0 &&
            getComputedStyle(element).visibility !== 'hidden';
        const buttons = [...document.querySelectorAll(
            'button, [role="button"], [data-testid*="send"], [data-icon="send"]'
        )].filter((button) => {
            const label = [
                button.innerText,
                button.getAttribute('aria-label'),
                button.getAttribute('title'),
                button.getAttribute('data-testid'),
                button.getAttribute('data-icon'),
                ...[...button.querySelectorAll('[data-icon]')]
                    .map((icon) => icon.getAttribute('data-icon'))
            ].filter(Boolean).join(' ').toLowerCase();
            return visible(button) && /send|enviar/.test(label);
        });
        for (const button of buttons) {
            let node = button;
            for (let depth = 0; node && node !== document.body && depth < 8; depth++) {
                if (visible(node) && node.querySelector(
                    'textarea, [contenteditable="true"], [data-placeholder]'
                )) return node;
                node = node.parentElement;
            }
        }
        return false;
        """
    )
    if send_container and _visible(send_container):
        return send_container
    return False


def _media_preview_diagnostics(driver: webdriver.Chrome) -> str:
    diagnostics = driver.execute_script(
        """
        const visible = (element) =>
            element.getClientRects().length > 0 &&
            getComputedStyle(element).visibility !== 'hidden';
        const describe = (element) => ({
            tag: element.tagName.toLowerCase(),
            aria: element.getAttribute('aria-label'),
            placeholder: element.getAttribute('placeholder'),
            dataPlaceholder: element.getAttribute('data-placeholder'),
            testId: element.getAttribute('data-testid'),
            accept: element.getAttribute('accept'),
            files: element.files ? [...element.files].map((file) => file.name) : [],
            visible: visible(element)
        });
        const previews = [...document.querySelectorAll(
            '[data-testid*="media-editor"], [data-testid*="media-preview"], ' +
            '[data-testid="image-thumb"]'
        )].filter(visible).map(describe);
        const canvas = document.querySelector(
            '[data-testid="media-editor-canvas"]'
        );
        const editorAncestors = [];
        for (let node = canvas, depth = 0;
             node && node !== document.body && depth < 8;
             node = node.parentElement, depth++) {
            editorAncestors.push({
                tag: node.tagName.toLowerCase(),
                testId: node.getAttribute('data-testid'),
                role: node.getAttribute('role'),
                aria: node.getAttribute('aria-label'),
                className: typeof node.className === 'string'
                    ? node.className.slice(0, 120) : ''
            });
        }
        const sendControls = [...document.querySelectorAll(
            'button, [role="button"], [data-testid], [data-icon]'
        )].filter(visible).filter((element) => {
            const label = [
                element.innerText,
                element.getAttribute('aria-label'),
                element.getAttribute('title'),
                element.getAttribute('data-testid'),
                element.getAttribute('data-icon')
            ].filter(Boolean).join(' ').toLowerCase();
            return /send|enviar/.test(label);
        }).slice(0, 20).map(describe);
        return JSON.stringify({
            dialogs: [...document.querySelectorAll('[role="dialog"]')]
                .filter(visible).length,
            previews,
            editorAncestors,
            sendControls,
            fileInputs: [...document.querySelectorAll('input[type="file"]')]
                .map(describe),
            editors: [...document.querySelectorAll(
                'textarea, [contenteditable="true"], [data-placeholder]'
            )].filter(visible).slice(-8).map(describe),
            buttons: [...document.querySelectorAll(
                'button, [role="button"], [data-testid*="send"], [data-icon="send"]'
            )].filter(visible).slice(-8).map(describe)
        });
        """
    )
    return str(diagnostics)


def _send_button(dialog: WebElement) -> WebElement | bool:
    for button in dialog.find_elements(
        By.CSS_SELECTOR,
        "button, [role='button'], [data-testid*='send'], [data-icon='send']",
    ):
        icon_labels = " ".join(
            icon.get_attribute("data-icon") or ""
            for icon in button.find_elements(By.CSS_SELECTOR, "[data-icon]")
        )
        label = " ".join(
            (
                button.text or "",
                button.get_attribute("aria-label") or "",
                button.get_attribute("title") or "",
                button.get_attribute("data-testid") or "",
                button.get_attribute("data-icon") or "",
                icon_labels,
            )
        ).casefold()
        if any(word in label for word in ("send", "enviar")) and _visible(button):
            return button
    return False


def _media_menu_item(driver: webdriver.Chrome) -> WebElement | bool:
    photo_video_labels = {
        "fotos e vídeos",
        "fotos e videos",
        "photos & videos",
        "photos and videos",
    }
    for element in driver.find_elements(
        By.XPATH,
        "//*[@role='menuitem'] | //*[@role='button'] | //button | //div",
    ):
        label = " ".join(
            (
                element.text or "",
                element.get_attribute("aria-label") or "",
                element.get_attribute("title") or "",
            )
        ).strip().casefold()
        if _visible(element) and label in photo_video_labels:
            return driver.execute_script(
                """
                return arguments[0].closest(
                    '[role="menuitem"], [role="button"], button'
                ) || arguments[0];
                """,
                element,
            )
    return False


def _click_media_menu_item(driver: webdriver.Chrome) -> None:
    for _attempt in range(3):
        try:
            item = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
                _media_menu_item
            )
            if not item:
                raise TimeoutException("A opção Fotos e vídeos não apareceu.")
            driver.execute_script("arguments[0].click()", item)
            return
        except StaleElementReferenceException:
            continue
    raise TimeoutException(
        "O menu de mídia foi atualizado repetidamente antes de selecionar "
        "Fotos e vídeos."
    )


def _send_images_with_caption(
    driver: webdriver.Chrome,
    images: list[Path],
    caption: str,
) -> None:
    if not images:
        raise ValueError("É necessário informar ao menos uma imagem.")

    stage = "abrir o menu de anexos"
    before_upload = ""
    selected_input = ""
    try:
        attachment_button = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            _attachment_button
        )
        if not attachment_button:
            raise TimeoutException("O botão de anexos não apareceu.")
        driver.execute_script("arguments[0].click()", attachment_button)

        stage = "selecionar Fotos e vídeos"
        _click_media_menu_item(driver)

        stage = "localizar o seletor de fotos"
        file_input = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            lambda browser: _find_file_input(browser, images_only=True)
        )
        supports_multiple = bool(file_input.get_property("multiple"))
        if len(images) > 1 and not supports_multiple:
            raise RuntimeError(
                "O seletor de fotos do WhatsApp não aceita múltiplos arquivos. "
                "Nenhuma imagem foi anexada; use um seletor que permita enviar "
                "o álbum com a legenda em uma única mensagem."
            )
        before_upload = _media_preview_diagnostics(driver)
        selected_input = str(
            driver.execute_script(
                """
                return JSON.stringify({
                    accept: arguments[0].getAttribute('accept'),
                    testId: arguments[0].getAttribute('data-testid'),
                    multiple: arguments[0].multiple
                });
                """,
                file_input,
            )
        )
        stage = "anexar as fotos"
        file_input.send_keys(
            "\n".join(str(image.resolve()) for image in images)
        )
        stage = "aguardar a prévia das imagens"
        dialog = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(_media_dialog)
        if not dialog:
            raise TimeoutException("A prévia das imagens não apareceu.")

        if caption:
            stage = "localizar o campo da descrição"

            def caption_box_ready(browser: webdriver.Chrome) -> WebElement | bool:
                active_dialog = _media_dialog(browser)
                if not active_dialog:
                    return False
                return next(
                    (
                        element
                        for element in active_dialog.find_elements(
                            By.CSS_SELECTOR,
                            "[contenteditable='true'], textarea, [data-placeholder]",
                        )
                        if _visible(element)
                        and element.get_attribute("data-testid")
                        != "conversation-compose-box-input"
                    ),
                    False,
                )

            caption_box = WebDriverWait(
                driver, CAPTION_FIELD_TIMEOUT_SECONDS
            ).until(caption_box_ready)
            if not caption_box:
                raise TimeoutException("O campo da descrição não apareceu.")
            stage = "preencher a descrição"
            caption_lines = (
                caption.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            )
            for index, line in enumerate(caption_lines):
                if index:
                    caption_box.send_keys(Keys.SHIFT, Keys.ENTER)
                if line:
                    caption_box.send_keys(line)

        stage = "localizar o botão Enviar"
        send_button = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            lambda browser: (
                _send_button(active_dialog)
                if (active_dialog := _media_dialog(browser))
                else False
            )
        )
        if not send_button:
            raise TimeoutException("O botão de envio das imagens não apareceu.")
        stage = "clicar no botão Enviar"
        driver.execute_script("arguments[0].click()", send_button)
        stage = "confirmar o envio das imagens"
        WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            lambda browser: not _media_dialog(browser)
        )
    except TimeoutException as error:
        details = (
            f" Seletor de arquivo usado: {selected_input}. "
            f"Estado antes do anexo: {before_upload}. "
            f"Estado após o anexo: {_media_preview_diagnostics(driver)}."
            if stage
            in {
                "aguardar a prévia das imagens",
                "localizar o campo da descrição",
                "preencher a descrição",
                "localizar o botão Enviar",
                "clicar no botão Enviar",
                "anexar as fotos",
            }
            else ""
        )
        raise RuntimeError(
            f"Não foi possível {stage} no WhatsApp Web. "
            f"Detalhe: {error}.{details}"
        ) from error
    except WebDriverException as error:
        raise RuntimeError(
            f"Falha do navegador durante a etapa '{stage}' ao enviar imagens "
            "pelo WhatsApp Web: "
            f"{error.msg.splitlines()[0]}"
        ) from error


def _send_media(
    driver: webdriver.Chrome,
    images: list[Path],
    caption: str,
) -> None:
    if not images:
        raise ValueError("É necessário informar ao menos uma imagem.")
    _send_images_with_caption(driver, images, caption)


def _attachment_button(driver: webdriver.Chrome) -> WebElement | bool:
    for button in driver.find_elements(By.CSS_SELECTOR, "button"):
        label = " ".join(
            (
                button.text or "",
                button.get_attribute("aria-label") or "",
                button.get_attribute("title") or "",
                button.get_attribute("data-testid") or "",
            )
        ).casefold()
        if _visible(button) and any(
            word in label for word in ("attach", "anexar", "clip", "anexo")
        ):
            return button
    return False


def _document_menu_item(driver: webdriver.Chrome) -> WebElement | bool:
    for element in driver.find_elements(
        By.XPATH,
        "//*[@role='menuitem'] | //*[normalize-space(.)='Documento' "
        "or normalize-space(.)='Document']",
    ):
        if _visible(element) and element.text.strip().casefold() in {
            "documento",
            "document",
        }:
            return element
    return False


def _send_document(driver: webdriver.Chrome, archive_path: Path) -> None:
    try:
        attach = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            _attachment_button
        )
        if not attach:
            raise TimeoutException("O botão de anexar arquivo não apareceu.")
        driver.execute_script("arguments[0].click()", attach)
        document_item = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            _document_menu_item
        )
        if not document_item:
            raise TimeoutException("A opção Documento não apareceu.")
        driver.execute_script("arguments[0].click()", document_item)
        file_input = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            lambda browser: _find_file_input(browser, images_only=False)
        )
        file_input.send_keys(str(archive_path.resolve()))
        dialog = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(_media_dialog)
        if not dialog:
            raise TimeoutException("A prévia do ZIP não apareceu.")
        send_button = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            lambda browser: (
                _send_button(active_dialog)
                if (active_dialog := _media_dialog(browser))
                else False
            )
        )
        if not send_button:
            raise TimeoutException("O botão de envio do ZIP não apareceu.")
        driver.execute_script("arguments[0].click()", send_button)
        WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(
            lambda browser: not _media_dialog(browser)
        )
    except TimeoutException as error:
        raise RuntimeError(
            "Não foi possível preparar ou confirmar o envio do ZIP no WhatsApp Web."
        ) from error
    except WebDriverException as error:
        raise RuntimeError(
            "Falha do navegador ao enviar o ZIP pelo WhatsApp Web: "
            f"{error.msg.splitlines()[0]}"
        ) from error


def _create_browser(profile_dir: Path | None) -> webdriver.Chrome:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise RuntimeError("A variável LOCALAPPDATA não está definida.")
    user_data_dir = profile_dir or (
        Path(local_app_data) / "SalesforceAutomation" / "WhatsAppProfile"
    )
    user_data_dir = user_data_dir.expanduser().resolve()
    user_data_dir.mkdir(parents=True, exist_ok=True)
    options = Options()
    options.add_argument(f"--user-data-dir={user_data_dir}")
    try:
        return webdriver.Chrome(options=options)
    except WebDriverException as error:
        raise RuntimeError(
            "Não foi possível iniciar o Chrome para o WhatsApp. Feche outras "
            "janelas que estejam usando o perfil dedicado."
        ) from error


def enviar_respostas(
    output_dir: Path,
    report_rows: dict[str, dict[str, str]],
    group_name: str,
    profile_dir: Path | None = None,
    input_fn: Callable[[str], str] = input,
) -> None:
    group_name = group_name.strip()
    if not group_name:
        raise ValueError("Informe o nome do grupo de WhatsApp.")

    archive_path = create_responses_archive(output_dir)
    archive_hash = _sha256_file(archive_path)
    state_path = _state_path(output_dir)
    state = _load_state(state_path)
    sent_notes = state["notes"]
    pending_notes: list[tuple[str, str, list[Path], str]] = []

    for note, row in report_rows.items():
        if row.get("status") not in SUCCESSFUL_SALESFORCE_STATUSES:
            continue
        note_dir = output_dir / note
        images = (
            sorted(
                path
                for path in note_dir.iterdir()
                if path.is_file() and path.suffix.casefold() in IMAGE_EXTENSIONS
            )
            if note_dir.is_dir()
            else []
        )
        if not images:
            print(f"WhatsApp: nota {note} ignorada (sem imagens).")
            continue
        message = _note_message(note_dir / "descricao.txt", row)
        fingerprint = _note_fingerprint(note, message, images)
        if sent_notes.get(note) != fingerprint:
            pending_notes.append((note, message, images, fingerprint))

    archive_pending = state.get("archive_sha256") != archive_hash
    if not pending_notes and not archive_pending:
        print("WhatsApp: notas e ZIP já enviados; nenhum envio duplicado.")
        return

    driver = _create_browser(profile_dir)
    authenticated = False
    try:
        driver.get("https://web.whatsapp.com")
        _wait_for_search_input(driver)
        authenticated = True
        _open_group(driver, group_name)
        if not _confirm(
            f"Confirme que o grupo aberto é '{group_name}' e autorize os envios",
            input_fn,
        ):
            print("WhatsApp: envios cancelados.")
            return

        for note, message, images, fingerprint in pending_notes:
            _send_media(driver, images, message)
            sent_notes[note] = fingerprint
            _save_state(state_path, state)
            print(f"WhatsApp: nota {note} enviada.")

        if archive_pending:
            _send_document(driver, archive_path)
            state["archive_sha256"] = archive_hash
            _save_state(state_path, state)
            print("WhatsApp: ZIP de Respostas enviado.")
    except WebDriverException as error:
        raise RuntimeError(
            "Falha do navegador durante o envio pelo WhatsApp Web: "
            f"{error.msg.splitlines()[0]}"
        ) from error
    finally:
        if authenticated:
            driver.quit()
        else:
            driver.service.stop()
