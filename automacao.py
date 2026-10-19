import argparse
import csv
import os
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


DEFAULT_WORKBOOK = Path(r"C:\automação salesforce\Geoex_Metropolitana.xlsx")
DEFAULT_OUTPUT = Path(r"C:\automação salesforce\Respostas")
DEFAULT_SHEET = "Notas sem evidências"
DEFAULT_COLUMN = "Nota"
DEFAULT_SALESFORCE_URL = "https://wfmneoenergia.my.salesforce.com/"
DEFAULT_WHATSAPP_GROUP = "Pendências Geoex - Metropolitana"
REPORT_NAME = "relatorio_processamento.csv"
LEGACY_REPORT_FIELDS = (
    "nota",
    "status",
    "motivo_nao_conclusao",
    "arquivos_baixados",
    "erro",
)
METADATA_FIELDS = (
    "texto_code_acao",
    "codigo_acao",
    "data_encerramento",
    "centro_trabalho",
)
REPORT_FIELDS = (
    "nota",
    *METADATA_FIELDS,
    "UTD",
    "status",
    "motivo_nao_conclusao",
    "arquivos_baixados",
    "erro",
)
SUCCESSFUL_SALESFORCE_STATUSES = frozenset({"Concluída", "Sem imagens"})
SOURCE_METADATA_COLUMNS = {
    "texto_code_acao": "Texto code ação",
    "codigo_acao": "Código de ação",
    "data_encerramento": "Data encermto.",
    "centro_trabalho": "CenTrab respon.",
    "UTD": "UTD",
}


def normalize_note(value: Any, row_number: int) -> str | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, bool):
        raise ValueError(f"Linha {row_number}: valor inválido para número da nota.")
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not value.is_integer():
            raise ValueError(f"Linha {row_number}: número da nota não é inteiro.")
        return str(int(value))
    if isinstance(value, str):
        note = value.strip()
        if not note.isascii() or not note.isdigit():
            raise ValueError(
                f"Linha {row_number}: a nota deve conter somente algarismos."
            )
        return note
    raise ValueError(f"Linha {row_number}: tipo de dado inválido para número da nota.")


def load_notes(workbook_path: Path, sheet_name: str, column_name: str) -> list[str]:
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Planilha não encontrada: {workbook_path}")

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            available = ", ".join(workbook.sheetnames)
            raise ValueError(
                f"Aba '{sheet_name}' não encontrada. Abas disponíveis: {available}"
            )

        sheet = workbook[sheet_name]
        header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        matching_columns = [
            index
            for index, value in enumerate(header)
            if isinstance(value, str) and value.strip().casefold() == column_name.casefold()
        ]
        if not matching_columns:
            available = ", ".join(str(value) for value in header if value is not None)
            raise ValueError(
                f"Coluna '{column_name}' não encontrada na aba '{sheet_name}'. "
                f"Colunas disponíveis: {available}"
            )
        if len(matching_columns) > 1:
            raise ValueError(
                f"A coluna '{column_name}' aparece mais de uma vez na aba '{sheet_name}'."
            )

        column_index = matching_columns[0]
        notes: list[str] = []
        seen: set[str] = set()
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            note = normalize_note(row[column_index], row_number)
            if note is not None and note not in seen:
                notes.append(note)
                seen.add(note)

        if not notes:
            raise ValueError(
                f"Nenhuma nota preenchida foi encontrada na coluna '{column_name}'."
            )
        return notes
    finally:
        workbook.close()


def _format_metadata_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()


def load_note_metadata(
    workbook_path: Path,
    sheet_name: str,
    column_name: str,
    notes: list[str],
) -> dict[str, dict[str, str]]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        sheet = workbook[sheet_name]
        header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())

        def find_columns(name: str) -> list[int]:
            return [
                index
                for index, value in enumerate(header)
                if isinstance(value, str)
                and value.strip().casefold() == name.casefold()
            ]

        note_columns = find_columns(column_name)
        if len(note_columns) != 1:
            raise ValueError(
                f"A coluna de nota '{column_name}' precisa aparecer exatamente "
                f"uma vez na aba '{sheet_name}'."
            )

        metadata_columns: dict[str, list[int]] = {}
        for field, name in SOURCE_METADATA_COLUMNS.items():
            indices = find_columns(name)
            if not indices:
                raise ValueError(
                    f"Coluna obrigatória '{name}' não encontrada na aba "
                    f"'{sheet_name}'."
                )
            metadata_columns[field] = indices

        note_set = set(notes)
        metadata: dict[str, dict[str, str]] = {}
        note_index = note_columns[0]
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            note = normalize_note(row[note_index], row_number)
            if note is None or note not in note_set:
                continue

            values: dict[str, str] = {}
            for field, indices in metadata_columns.items():
                column_values = [
                    _format_metadata_value(row[index]) for index in indices
                ]
                if len(set(column_values)) > 1:
                    raise ValueError(
                        f"As colunas duplicadas 'Texto code ação' têm valores "
                        f"diferentes na linha {row_number}."
                    )
                values[field] = column_values[0]

            if note in metadata and metadata[note] != values:
                raise ValueError(
                    f"A nota {note} aparece com descrições diferentes na planilha."
                )
            metadata[note] = values

        missing = note_set - metadata.keys()
        if missing:
            raise ValueError(
                f"Não foi possível carregar a descrição das notas: "
                f"{', '.join(sorted(missing))}"
            )
        return metadata
    finally:
        workbook.close()


def write_note_descriptions(
    output_dir: Path,
    notes: list[str],
    metadata: dict[str, dict[str, str]],
) -> None:
    labels = {
        "texto_code_acao": "Texto code ação",
        "codigo_acao": "Código de ação",
        "data_encerramento": "Data de encerramento",
        "centro_trabalho": "Centro de trabalho responsável",
    }
    for note in notes:
        description_path = output_dir / note / "descricao.txt"
        with description_path.open("w", encoding="utf-8", newline="\n") as file:
            file.write(f"Nota: {note}\n")
            for field in METADATA_FIELDS:
                file.write(f"{labels[field]}: {metadata[note][field]}\n")


def read_existing_report(report_path: Path) -> dict[str, dict[str, str]]:
    if not report_path.exists():
        return {}

    with report_path.open("r", encoding="utf-8-sig", newline="") as report_file:
        reader = csv.DictReader(report_file, delimiter=";")
        if reader.fieldnames is None or not set(LEGACY_REPORT_FIELDS).issubset(
            reader.fieldnames
        ):
            raise ValueError(
                f"O arquivo '{report_path}' existe, mas não tem o formato esperado. "
                "Renomeie-o antes de gerar um novo relatório."
            )
        return {
            row["nota"]: {field: row.get(field, "") or "" for field in REPORT_FIELDS}
            for row in reader
            if row.get("nota")
        }


def write_report(
    report_path: Path,
    notes: list[str],
    updates: dict[str, dict[str, str]] | None = None,
) -> None:
    previous_rows = read_existing_report(report_path)
    updates = updates or {}
    all_notes = list(dict.fromkeys([*previous_rows, *notes, *updates]))
    report_path.parent.mkdir(parents=True, exist_ok=True)

    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{report_path.stem}_",
        suffix=".tmp",
        dir=report_path.parent,
    )
    try:
        with os.fdopen(
            file_descriptor, "w", encoding="utf-8-sig", newline=""
        ) as report_file:
            writer = csv.DictWriter(
                report_file,
                fieldnames=REPORT_FIELDS,
                delimiter=";",
                lineterminator="\n",
            )
            writer.writeheader()
            for note in all_notes:
                row = {field: "" for field in REPORT_FIELDS}
                row.update({"nota": note, "status": "Pendente"})
                row.update(previous_rows.get(note, {}))
                row.update(updates.get(note, {}))
                writer.writerow(row)
        try:
            os.replace(temporary_name, report_path)
        except PermissionError as error:
            raise PermissionError(
                f"Não foi possível atualizar o relatório '{report_path}'. "
                "Feche-o no Excel ou em qualquer outro programa que o esteja "
                "usando e execute novamente."
            ) from error
    finally:
        if os.path.exists(temporary_name):
            os.remove(temporary_name)


def prepare_report(
    report_path: Path,
    notes: list[str],
    metadata: dict[str, dict[str, str]],
) -> dict[str, dict[str, str]]:
    write_report(report_path, notes, metadata)
    return read_existing_report(report_path)


def select_notes_to_process(
    notes: list[str],
    existing_rows: dict[str, dict[str, str]],
    limit: int,
) -> list[str]:
    successful_notes = {
        note
        for note, row in existing_rows.items()
        if row.get("status") in SUCCESSFUL_SALESFORCE_STATUSES
    }
    return [note for note in notes if note not in successful_notes][:limit]


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Valida a planilha e prepara uma fila de notas e um relatório "
            "para a automação Salesforce."
        )
    )
    parser.add_argument("--planilha", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--aba", default=DEFAULT_SHEET)
    parser.add_argument("--coluna", default=DEFAULT_COLUMN)
    parser.add_argument("--saida", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument(
        "--consultar-salesforce",
        action="store_true",
        help="Consulta notas no Salesforce e baixa os arquivos dos compromissos.",
    )
    parser.add_argument(
        "--limite",
        type=int,
        default=1,
        help="Número máximo de notas para consultar no Salesforce (padrão: 1).",
    )
    parser.add_argument("--salesforce-url", default=DEFAULT_SALESFORCE_URL)
    parser.add_argument(
        "--perfil-chrome",
        type=Path,
        default=None,
        help=(
            "Diretório do perfil dedicado do Chrome "
            "(padrão: %%LOCALAPPDATA%%\\SalesforceAutomation\\ChromeProfile)."
        ),
    )
    parser.add_argument(
        "--enviar-whatsapp",
        action="store_true",
        help=(
            "Publica imagens e descrições no grupo do WhatsApp e envia um ZIP "
            "da pasta de respostas."
        ),
    )
    parser.add_argument(
        "--grupo-whatsapp",
        default=DEFAULT_WHATSAPP_GROUP,
        help=f"Nome exato do grupo de WhatsApp (padrão: {DEFAULT_WHATSAPP_GROUP}).",
    )
    parser.add_argument(
        "--perfil-whatsapp",
        type=Path,
        default=None,
        help=(
            "Diretório do perfil dedicado do Chrome para WhatsApp "
            "(padrão: %%LOCALAPPDATA%%\\SalesforceAutomation\\WhatsAppProfile)."
        ),
    )
    return parser.parse_args(argv)


def main(
    argv: list[str] | None = None,
    *,
    input_fn: Callable[[str], str] = input,
    output_fn: Callable[[str], None] = print,
    error_fn: Callable[[str], None] | None = None,
) -> int:
    args = parse_args(argv)
    try:
        notes = load_notes(args.planilha, args.aba, args.coluna)
        metadata = load_note_metadata(
            args.planilha,
            args.aba,
            args.coluna,
            notes,
        )
        args.saida.mkdir(parents=True, exist_ok=True)
        report_path = args.saida / REPORT_NAME
        previous_rows = read_existing_report(report_path)
        new_notes = [note for note in notes if note not in previous_rows]
        notes_to_initialize = [
            note
            for note in notes
            if note not in previous_rows
            or not (args.saida / note / "descricao.txt").is_file()
        ]
        for note in notes:
            (args.saida / note).mkdir(exist_ok=True)
        metadata_to_initialize = {
            note: metadata[note] for note in notes_to_initialize
        }
        write_note_descriptions(args.saida, notes_to_initialize, metadata)
        existing_rows = prepare_report(
            report_path,
            notes,
            metadata_to_initialize,
        )
        if args.consultar_salesforce:
            if args.limite < 1:
                raise ValueError("--limite deve ser maior que zero.")
            notes_to_process = select_notes_to_process(
                [*new_notes, *(note for note in notes if note not in new_notes)],
                existing_rows,
                args.limite,
            )
            from salesforce import ResultadoNota, consultar_notas

            result_updates: dict[str, dict[str, str]] = {}

            def save_result(resultado: ResultadoNota) -> None:
                row = existing_rows[resultado.nota].copy()
                row.update(
                    {
                        "status": resultado.status,
                        "motivo_nao_conclusao": resultado.motivo,
                        "arquivos_baixados": ", ".join(resultado.arquivos),
                        "erro": resultado.erro,
                    }
                )
                result_updates[resultado.nota] = row
                write_report(report_path, notes, result_updates)

            if notes_to_process:
                resultados = consultar_notas(
                    notes_to_process,
                    args.salesforce_url,
                    args.saida,
                    args.perfil_chrome,
                    on_result=save_result,
                    output_fn=output_fn,
                )
            else:
                resultados = []
            skipped_success_count = sum(
                existing_rows.get(note, {}).get("status")
                in SUCCESSFUL_SALESFORCE_STATUSES
                for note in notes
            )
        if args.enviar_whatsapp:
            existing_rows = read_existing_report(report_path)
            from whatsapp import enviar_respostas

            enviar_respostas(
                args.saida,
                existing_rows,
                args.grupo_whatsapp,
                args.perfil_whatsapp,
                input_fn=input_fn,
                output_fn=output_fn,
            )
    except (OSError, InvalidFileException, ValueError) as error:
        message = f"Erro: {error}"
        if error_fn is None:
            print(message, file=sys.stderr)
        else:
            error_fn(message)
        return 1
    except RuntimeError as error:
        message = f"Erro: {error}"
        if error_fn is None:
            print(message, file=sys.stderr)
        else:
            error_fn(message)
        return 1

    output_fn(f"Notas preparadas: {len(notes)}")
    output_fn(f"Novas notas adicionadas ao relatório: {len(new_notes)}")
    output_fn(f"Relatório: {report_path}")
    output_fn(f"Pastas por nota: {args.saida}")
    if args.consultar_salesforce:
        output_fn(f"Notas consultadas no Salesforce: {len(resultados)}")
        output_fn(f"Notas já concluídas ignoradas: {skipped_success_count}")
    else:
        output_fn("Nenhuma consulta ou alteração foi feita no Salesforce.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
