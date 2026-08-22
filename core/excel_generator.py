import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Side


class ExcelGenerator:

    def __init__(self, language):
        self.language = language

    def generate(
        self,
        students,
        assignments,
        practice_questions,
        questions_per_student,
        course,
        module,
        term,
        course_language
    ):

        """
        Generate the Excel assignment workbook.
        """

        # Create workbook and worksheets
        workbook = Workbook()

        workbook.properties.creator = "Ing. Lucio Buitrón"
        workbook.properties.lastModifiedBy = "Ing. Lucio Buitrón"
        workbook.properties.title = "Question Assignment Tool"
        workbook.properties.subject = "Question Assignment"
        workbook.properties.description = (
            "Assignment file generated with Question Assignment Tool v2.0"
        )
        workbook.properties.keywords = (
            "Question Assignment, Education, Assessment, Excel"
        )
        workbook.properties.category = "Education"

        worksheet = workbook.active
        worksheet.title = "Theoretical Questions"
        practice_worksheet = workbook.create_sheet("Practice Questions")

        header_fill = PatternFill(
            fill_type="solid",
            start_color="1F4E78"
        )

        light_blue_fill = PatternFill(
            fill_type="solid",
            start_color="EAF4FF"
        )

        header_font = Font(
            bold=True,
            color="FFFFFF"
        )

        header_alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        center_alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        left_alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )

        # =====================================================
        # Grade Reference - Top Section
        # =====================================================

        # Configure grade reference
        grade_reference = [
            ("EXCELLENT", 260),
            ("GOOD", 195),
            ("FAIR", 130),
            ("POOR", 65)
        ]

        # Grade Reference - Assignments

        grade_start_row = 1

        worksheet.cell(row=1, column=1).value = "Grade Reference"
        worksheet.cell(row=1, column=1).font = Font(bold=True, size=8)

        worksheet.cell(row=2, column=1).value = "Grade"
        worksheet.cell(row=2, column=2).value = "Points"

        for cell in worksheet[2]:
            if cell.column <= 2:
                cell.fill = header_fill
                cell.font = Font(
                    bold=True,
                    color="FFFFFF",
                    size=8
                )
                cell.alignment = header_alignment

        for index, (grade, points) in enumerate(
            grade_reference,
            start=3
        ):
            worksheet.cell(row=index, column=1).value = grade
            worksheet.cell(row=index, column=2).value = points
            worksheet.cell(row=index, column=1).alignment = center_alignment
            worksheet.cell(row=index, column=2).alignment = left_alignment
            worksheet.cell(row=index, column=1).border = thin_border
            worksheet.cell(row=index, column=2).border = thin_border

        for row in range(1, 7):
            for column in range(1, 3):
                worksheet.cell(
                    row=row,
                    column=column
                ).font = Font(
                    bold=(row in (1, 2)),
                    color="FFFFFF" if row == 2 else "000000",
                    size=8
                )        
        
        grade_note_row = 7

        worksheet.cell(
            row=grade_note_row,
            column=1
        ).value = self.language.get("grade_note")

        worksheet.merge_cells(
            start_row=grade_note_row,
            start_column=1,
            end_row=grade_note_row,
            end_column=5
        )

        worksheet.cell(
            row=grade_note_row,
            column=1
        ).alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )

        worksheet.cell(
            row=grade_note_row,
            column=1
        ).font = Font(bold=True, size=9)

        worksheet.cell(
            row=grade_note_row,
            column=1
        ).fill = light_blue_fill

        # Assignments table starts after Grade Reference
        worksheet.cell(row=10, column=1).value = "#"
        worksheet.cell(row=10, column=2).value = "Student"
        worksheet.cell(row=10, column=3).value = "Theoretical Question"
        worksheet.cell(row=10, column=4).value = "Grade"
        worksheet.cell(row=10, column=5).value = "Final Grade"
        worksheet.cell(row=10, column=6).value = "Final Grade ± Defense Evaluation"
        worksheet.cell(row=10, column=7).value = "Comment"

        # ================================
        # Practice Questions Section
        # ================================

        # Practice Questions table starts after Grade Reference
        practice_worksheet.cell(row=1, column=1).value = "#"
        practice_worksheet.cell(row=1, column=2).value = "Practice Question"
        practice_worksheet.cell(row=1, column=3).value = "Group"
        practice_worksheet.cell(row=1, column=4).value = "Student (Optional)"
        practice_worksheet.cell(row=1, column=5).value = "Grade (Only Reference)"
        practice_worksheet.cell(row=1, column=6).value = "Comment"

        # Create Practice Questions validations
        group_validation = DataValidation(
            type="list",
            formula1='"Group 1,Group 2,Group 3,Group 4,Group 5,Group 6,Group 7,Group 8,Group 9,Group 10"',
            allow_blank=True
        )

        practice_grade_validation = DataValidation(
            type="list",
            formula1='"Not Graded,Excellent,Good,Fair,Poor"',
            allow_blank=True
        )

        # Create hidden student list for dropdown
        student_list_sheet = workbook.create_sheet("_Lists")

        for index, student in enumerate(students, start=1):
            student_list_sheet.cell(row=index, column=1).value = student

        student_list_sheet.sheet_state = "hidden"

        student_validation = DataValidation(
            type="list",
            formula1=f"'_Lists'!$A$1:$A${len(students)}",
            allow_blank=True
        )

        practice_worksheet.add_data_validation(group_validation)
        practice_worksheet.add_data_validation(student_validation)
        practice_worksheet.add_data_validation(practice_grade_validation)

        # Populate Practice Questions
        for practice_question in practice_questions:

            row = practice_worksheet.max_row + 1

            practice_worksheet.cell(row=row, column=1).value = row - 1
            practice_worksheet.cell(row=row, column=2).value = practice_question
            practice_worksheet.cell(row=row, column=3).value = ""
            practice_worksheet.cell(row=row, column=4).value = ""
            practice_worksheet.cell(row=row, column=5).value = ""
            practice_worksheet.cell(row=row, column=6).value = ""

            group_validation.add(
                practice_worksheet.cell(row=row, column=3)
            )

            student_validation.add(
                practice_worksheet.cell(row=row, column=4)
            )

            practice_grade_validation.add(
                practice_worksheet.cell(row=row, column=5)
            )

        # Practice Questions formatting
        for row in practice_worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border

        for cell in practice_worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        for row in range(2, practice_worksheet.max_row + 1):

            practice_worksheet.cell(row=row, column=1).alignment = center_alignment
            practice_worksheet.cell(row=row, column=2).alignment = left_alignment
            practice_worksheet.cell(row=row, column=3).alignment = center_alignment
            practice_worksheet.cell(row=row, column=4).alignment = left_alignment
            practice_worksheet.cell(row=row, column=5).alignment = center_alignment
            practice_worksheet.cell(row=row, column=6).alignment = left_alignment

            # Alternating row color
            if row % 2 == 0:
                for column in range(1, 7):
                    practice_worksheet.cell(
                        row=row,
                        column=column
                    ).fill = light_blue_fill

        practice_worksheet.column_dimensions["A"].width = 10
        practice_worksheet.column_dimensions["B"].width = 90
        practice_worksheet.column_dimensions["C"].width = 15
        practice_worksheet.column_dimensions["D"].width = 40
        practice_worksheet.column_dimensions["E"].width = 25
        practice_worksheet.column_dimensions["F"].width = 55
        
        for cell in worksheet[10]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        worksheet.cell(row=10, column=2).alignment = header_alignment
        worksheet.cell(row=10, column=3).alignment = header_alignment

        for column in range(5, 8):
            worksheet.cell(
                row=10,
                column=column
            ).fill = header_fill

            worksheet.cell(
                row=10,
                column=column
            ).font = header_font

            worksheet.cell(
                row=10,
                column=6
            ).alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
             
        grade_validation = DataValidation(
            type="list",
            formula1='"Not Graded,Excellent,Good,Fair,Poor"',
            allow_blank=True
        )

        worksheet.add_data_validation(grade_validation)

        # Populate Theoretical Questions
        row = 11
        student_number = 1

        for student, assigned_questions in assignments.items():

            start_row = row

            for question in assigned_questions:

                worksheet.cell(row=row, column=3).value = question

                worksheet.cell(
                    row=row,
                    column=3
                ).alignment = left_alignment

                worksheet.cell(row=row, column=4).value = ""
                worksheet.cell(
                    row=row,
                    column=4
                ).alignment = center_alignment

                grade_validation.add(
                    worksheet.cell(row=row, column=4)
                )
                worksheet.cell(row=row, column=5).value = ""
                worksheet.cell(row=row, column=6).value = ""
                worksheet.cell(
                    row=row,
                    column=6
                ).alignment = center_alignment
                worksheet.cell(row=row, column=7).value = ""
                worksheet.cell(
                    row=row,
                    column=7
                ).alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=True
                )
                row += 1
            end_row = row - 1

            worksheet.cell(row=start_row, column=1).value = student_number
            worksheet.cell(
                row=start_row,
                column=1
            ).alignment = center_alignment

            worksheet.cell(row=start_row, column=2).value = student
            worksheet.cell(
                row=start_row,
                column=2
            ).alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=True
            )

            worksheet.cell(
                row=start_row,
                column=5
            ).value = (
                f'=IFERROR(('
                f'COUNTIF(D{start_row}:D{end_row},"EXCELLENT")*260+'
                f'COUNTIF(D{start_row}:D{end_row},"GOOD")*195+'
                f'COUNTIF(D{start_row}:D{end_row},"FAIR")*130+'
                f'COUNTIF(D{start_row}:D{end_row},"POOR")*65'
                f')/('
                f'COUNTIF(D{start_row}:D{end_row},"EXCELLENT")+'
                f'COUNTIF(D{start_row}:D{end_row},"GOOD")+'
                f'COUNTIF(D{start_row}:D{end_row},"FAIR")+'
                f'COUNTIF(D{start_row}:D{end_row},"POOR")'
                f'),"")'
            )

            worksheet.cell(
                row=start_row,
                column=5
            ).alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            if end_row > start_row:
                worksheet.merge_cells(
                    start_row=start_row,
                    start_column=1,
                    end_row=end_row,
                    end_column=1
                )

                worksheet.merge_cells(
                    start_row=start_row,
                    start_column=2,
                    end_row=end_row,
                    end_column=2
                )

                worksheet.merge_cells(
                    start_row=start_row,
                    start_column=5,
                    end_row=end_row,
                    end_column=5
                )

                worksheet.merge_cells(
                    start_row=start_row,
                    start_column=6,
                    end_row=end_row,
                    end_column=6
                )

                if student_number % 2 == 1:
                    for r in range(start_row, end_row + 1):
                        for c in range(1, 8):   # Columnas A-E
                            worksheet.cell(row=r, column=c).fill = light_blue_fill

            student_number += 1

        worksheet.column_dimensions["A"].width = 10
        worksheet.column_dimensions["B"].width = 45
        worksheet.column_dimensions["C"].width = 100
        worksheet.column_dimensions["D"].width = 13
        worksheet.column_dimensions["E"].width = 14
        worksheet.column_dimensions["F"].width = 14
        worksheet.column_dimensions["G"].width = 50

        # Apply final worksheet formatting
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border

            worksheet.cell(
                row=grade_note_row,
                column=1
            ).alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=True
            )

            worksheet.cell(
                row=10,
                column=2
            ).alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            worksheet.cell(
                row=10,
                column=3
            ).alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # =====================================================
        # Save Excel
        # =====================================================

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        course = course.replace(" ", "")
        filename = (
            f"{course}_"
            f"{module}_"
            f"Term{term}_"
            f"{course_language}_"
            f"{timestamp}.xlsx"
        )

        output_folder = "output"
        os.makedirs(output_folder, exist_ok=True)
        filepath = os.path.join(output_folder, filename)
        filepath = os.path.abspath(filepath)

        workbook.save(filepath)
        return filepath
