import logging
from enum import Enum
from typing import Any

import click
import click_log
import gspread
from gspread import NoValidUrlKeyFound, SpreadsheetNotFound

logger = logging.getLogger(__name__)
click_log.basic_config(logger)

# Cell containing amount of total lectures ("Total de aulas no semestre:") in A1 notation
TOTAL_LECTURES_CELL = "A2:H2"
# Index of row containing key cells (Matricula, Aluno, Faltas etc.)
HEAD_ROW_INDEX = 3


class ColumnKey(str, Enum):
    ID = "Matricula"
    STUDENT = "Aluno"
    ABSCENCES = "Faltas"
    P1 = "P1"
    P2 = "P2"
    P3 = "P3"
    SITUATION = "Situação"
    NAF = "Nota para Aprovação Final"

    def __str__(self) -> str:
        return self.value


@click.command(
    help="""Program to calculate student approval situations based on their test scores and abscences.

SHEET_ID may be the URL, key or name of the spreadsheet."""
)
@click_log.simple_verbosity_option(logger)
@click.argument("spreadsheet_id")
def main(spreadsheet_id: str):
    logger.info("Initializing Google Sheets client...")
    client = gspread.oauth()

    logger.info("Searching spreadsheet...")
    try:
        sheet = client.open_by_url(spreadsheet_id)
    except (NoValidUrlKeyFound, SpreadsheetNotFound):
        try:
            sheet = client.open_by_key(spreadsheet_id)
        except (NoValidUrlKeyFound, SpreadsheetNotFound):
            try:
                sheet = client.open(spreadsheet_id)
            except SpreadsheetNotFound:
                logger.error(
                    "Could not open the spreadsheet. Check the given URL/key/name and try again."
                )
                return

    logger.info(f"Found worksheets {sheet.worksheets()}, selecting first one...")
    worksheet = sheet.sheet1

    logger.info("Reading total lectures amount...")
    total_lectures_cell = worksheet.acell(TOTAL_LECTURES_CELL)
    if total_lectures_cell.value is None:
        logger.error(
            f"Total lectures amount cell ({TOTAL_LECTURES_CELL}) does not exist!"
        )
        return
    total_lectures = int(total_lectures_cell.value.split(" ")[-1])
    logger.info(f"Total lectures: {total_lectures}")

    logger.info("Reading student records...")
    student_records = worksheet.get_all_records(head=HEAD_ROW_INDEX, empty2zero=True)

    situation_cell = worksheet.find(ColumnKey.SITUATION, in_row=HEAD_ROW_INDEX)
    if situation_cell is None:
        logger.error(
            f"Could not find student situation column! ({ColumnKey.SITUATION})"
        )
        return
    nafs_cell = worksheet.find(ColumnKey.NAF, in_row=HEAD_ROW_INDEX)
    if nafs_cell is None:
        logger.error(f"Could not find student NAF column! ({ColumnKey.NAF})")
        return

    situation_col = worksheet.range(
        situation_cell.row + 1,
        situation_cell.col,
        situation_cell.row + len(student_records) + 1,
        situation_cell.col,
    )
    nafs_col = worksheet.range(
        nafs_cell.row + 1,
        nafs_cell.col,
        nafs_cell.row + len(student_records) + 1,
        nafs_cell.col,
    )

    for i, record in enumerate(student_records):
        record: dict[str, Any]
        student_id = record[ColumnKey.ID]
        student_name = record[ColumnKey.STUDENT]
        logger.info(f"Computing situation for student {student_name} ({student_id})...")

        (situation, naf) = compute_student_situation(
            abscences=record[ColumnKey.ABSCENCES],
            p1=record[ColumnKey.P1],
            p2=record[ColumnKey.P2],
            p3=record[ColumnKey.P3],
            total_lectures=total_lectures,
        )
        logger.info(f"Situation: {situation}, NAF: {naf}")

        situation_col[i].value = situation
        nafs_col[i].value = naf

    logger.info("Writing results to worksheet...")
    worksheet.update_cells(situation_col)
    worksheet.update_cells(nafs_col)

    logger.info("All done!")


class StudentSituation(str, Enum):
    REPROVADO_POR_NOTA = "Reprovado por Nota"
    REPROVADO_POR_FALTA = "Reprovado por Falta"
    EXAME_FINAL = "Exame Final"
    APROVADO = "Aprovado"

    def __str__(self) -> str:
        return self.value


def compute_student_situation(
    abscences: int,
    p1: int,
    p2: int,
    p3: int,
    total_lectures: int,
) -> tuple[StudentSituation, int]:
    """Computes a student's situation and NAF based on their test scores and abscences.

    The student's situation is determined by the average of their test scores,
    following this table:

        average < 5      -> REPROVADO_POR_NOTA (failed by score)
        5 <= average < 7 -> EXAME_FINAL (final exam)
        average >= 7     -> APROVADO (approved)

    If the number of abscences exceeds 25% of the total amount of lectures,
    the student will have the REPROVADO_POR_FALTA (failed by abscences) situation.

    If the student's situation is "Exame Final", their "Nota para Aprovação Final" (NAF)
    will be calculated by ceil(10 - average).

    Args:
        abscences: Number of abscences.
        p1: First test score.
        p2: Second test score.
        p3: Third test score.
        total_lectures: Total amount of lectures in the semester.

    Returns:
        A tuple containing the student's situation and their NAF.

        If the student's situation is not EXAME_FINAL, the NAF value will be 0.
    """

    from math import ceil

    # Note: Test scores in the input data provided are in the scale of 0-100.
    # I assume this is a mistake in the problem description or input data,
    # so here the average is divided by 10 to match the 0-10 scale.
    average = (p1 + p2 + p3) / 3 / 10
    situation: StudentSituation
    naf: int = 0

    if abscences > (total_lectures / 4):
        situation = StudentSituation.REPROVADO_POR_FALTA
    elif average < 5:
        situation = StudentSituation.REPROVADO_POR_NOTA
    elif 5 <= average < 7:
        situation = StudentSituation.EXAME_FINAL
        naf = ceil(10 - average)
    else:
        situation = StudentSituation.APROVADO

    return (situation, naf)


if __name__ == "__main__":
    main()
