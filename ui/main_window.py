import json
import os
import customtkinter as ctk
import tkinter as tk

from core.language_manager import LanguageManager
from core.excel_generator import ExcelGenerator
from core.assignment_manager import AssignmentManager
from tkinter import messagebox
from core.path_utils import resource_path


class MainWindow(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.load_configuration()
        self.language = LanguageManager(
            self.config_data["general"]["application_language"]
        )
        self.create_menu()
        self.configure_window()
        self.create_widgets()

    def create_menu(self):

        self.menu_bar = tk.Menu(self)

        self.help_menu = tk.Menu(
            self.menu_bar,
            tearoff=0
        )

        self.menu_bar.add_cascade(
            label=self.language.get("help"),
            menu=self.help_menu
        )

        self.help_menu.add_command(
            label=self.language.get("instructions"),
            command=self.show_instructions
        )

        self.help_menu.add_separator()

        self.help_menu.add_command(
            label=self.language.get("about"),
            command=self.show_about
        )

        self.config(menu=self.menu_bar)

    def load_configuration(self):

        with open(resource_path("config.json"), "r", encoding="utf-8") as file:
            self.config_data = json.load(file)

    def configure_window(self):

        app = self.config_data["application"]
        window = self.config_data["window"]

        self.title(f'{self.language.get("title")} v{app["version"]}')

        screen_width = self.winfo_screenwidth()
        x = (screen_width - window["width"]) // 2
        y = 0

        self.geometry(
            f'{window["width"]}x{window["height"]}+{x}+{y}'
        )

        self.minsize(
            window["min_width"],
            window["min_height"]
        )

        self.resizable(
            window["resizable"],
            window["resizable"]
        )

    def create_widgets(self):

        lists = self.config_data["lists"]

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        for column in range(5):
            self.main_frame.grid_columnconfigure(column, weight=1)

        # =====================================================
        # Title
        # =====================================================

        self.lbl_title = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("title"),
            font=("Segoe UI", 24, "bold")
        )

        self.lbl_title.grid(
            row=0,
            column=0,
            columnspan=5,
            pady=(10, 20)
        )

        # =====================================================
        # Application Language
        # =====================================================

        self.lbl_application_language = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("application_language").upper()
        )

        self.lbl_application_language.grid(
            row=1,
            column=0,
            columnspan=5,
            pady=(0, 5)
        )

        self.cmb_application_language = ctk.CTkComboBox(
            self.main_frame,
            values=list(lists["application_languages"].keys()),
            command=self.change_language
        )

        language_code = self.config_data["general"]["application_language"]

        for name, code in lists["application_languages"].items():
            if code == language_code:
                self.cmb_application_language.set(name)
                break

        self.cmb_application_language.grid(
            row=2,
            column=2,
            padx=5,
            pady=(0, 20),
            sticky="ew"
        )

        # =====================================================
        # Course
        # =====================================================

        self.lbl_course = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("course").upper()
        )

        self.lbl_course.grid(row=3, column=0, padx=5, sticky="w")

        self.cmb_course = ctk.CTkComboBox(
            self.main_frame,
            values=lists["courses"]
        )

        self.cmb_course.set("Operating Systems 2")
        self.cmb_course.grid(row=4, column=0, padx=5, sticky="ew")

        # =====================================================
        # Module
        # =====================================================

        self.lbl_module = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("module").upper()
        )

        self.lbl_module.grid(row=3, column=1, padx=5, sticky="w")

        self.cmb_module = ctk.CTkComboBox(
            self.main_frame,
            values=lists["modules"]
        )

        self.cmb_module.grid(row=4, column=1, padx=5, sticky="ew")

        # =====================================================
        # Term
        # =====================================================

        self.lbl_term = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("term").upper()
        )

        self.lbl_term.grid(row=3, column=2, padx=5, sticky="w")

        self.cmb_term = ctk.CTkComboBox(
            self.main_frame,
            values=lists["terms"]
        )

        self.cmb_term.grid(row=4, column=2, padx=5, sticky="ew")

        # =====================================================
        # Year
        # =====================================================

        self.lbl_year = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("year").upper()
        )

        self.lbl_year.grid(row=3, column=3, padx=5, sticky="w")

        self.cmb_year = ctk.CTkComboBox(
            self.main_frame,
            values=[str(year) for year in range(2026, 2036)]
        )

        self.cmb_year.set("2026")

        self.cmb_year.grid(row=4, column=3, padx=5, sticky="ew")

        # =====================================================
        # Course Language
        # =====================================================

        self.lbl_course_language = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("course_language").upper()
        )

        self.lbl_course_language.grid(row=3, column=4, padx=5, sticky="w")

        self.cmb_course_language = ctk.CTkComboBox(
            self.main_frame,
            values=lists["course_languages"]
        )

        self.cmb_course_language.set("Spanish")
        self.cmb_course_language.grid(row=4, column=4, padx=5, sticky="ew")

        # =====================================================
        # Student List
        # =====================================================

        self.lbl_students = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("student_list").upper()
        )

        self.lbl_students.grid(
            row=5,
            column=0,
            columnspan=5,
            sticky="w",
            pady=(25, 5)
        )

        self.txt_students = ctk.CTkTextbox(
            self.main_frame,
            height=140
        )

        self.txt_students.grid(
            row=6,
            column=0,
            columnspan=5,
            sticky="nsew"
        )

        # =====================================================
        # Question Bank
        # =====================================================

        self.lbl_questions = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("question_bank").upper()
        )

        self.lbl_questions.grid(
            row=7,
            column=0,
            columnspan=5,
            sticky="w",
            pady=(25, 5)
        )

        self.txt_questions = ctk.CTkTextbox(
            self.main_frame,
            height=100
        )

        self.txt_questions.grid(
            row=8,
            column=0,
            columnspan=5,
            sticky="nsew"
        )

        # =====================================================
        # Questions per Student
        # =====================================================

        self.lbl_questions_per_student = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("questions_per_student").upper()
        )

        self.lbl_questions_per_student.grid(
            row=9,
            column=0,
            padx=5,
            pady=(20, 5),
            sticky="w"
        )

        self.cmb_questions_per_student = ctk.CTkComboBox(
            self.main_frame,
            values=lists["questions_per_student"]
        )

        self.cmb_questions_per_student.set("2")

        self.cmb_questions_per_student.grid(
            row=10,
            column=0,
            padx=5,
            sticky="ew"
        )

        # =====================================================
        # Practice Questions
        # =====================================================

        self.lbl_practice_questions = ctk.CTkLabel(
            self.main_frame,
            text=self.language.get("practice_questions").upper()
        )

        self.lbl_practice_questions.grid(
            row=11,
            column=0,
            columnspan=5,
            sticky="w",
            pady=(25, 5)
        )

        self.txt_practice_questions = ctk.CTkTextbox(
            self.main_frame,
            height=100
        )

        self.txt_practice_questions.grid(
            row=12,
            column=0,
            columnspan=5,
            sticky="nsew"
        )

        # =====================================================
        # Generate
        # =====================================================

        self.btn_generate = ctk.CTkButton(
            self.main_frame,
            text=self.language.get("generate_excel"),
            command=self.generate_excel
        )

        self.btn_generate.grid(
            row=13,
            column=0,
            columnspan=5,
            pady=25
        )

    
        # =====================================================
        # Exit
        # =====================================================

        self.btn_exit = ctk.CTkButton(
            self.main_frame,
            text=self.language.get("exit"),
            command=self.destroy
        )

        self.btn_exit.grid(
            row=16,
            column=0,
            columnspan=5,
            pady=(0, 10)
        )
    
    # =====================================================
    # Generate Excel
    # =====================================================
    def generate_excel(self):

        students = [
            student.strip()
            for student in self.txt_students.get("1.0", "end").splitlines()
            if student.strip()
        ]

        questions = [
            question.strip()
            for question in self.txt_questions.get("1.0", "end").splitlines()
            if question.strip()
        ]

        practice_questions = [
            question.strip()
            for question in self.txt_practice_questions.get("1.0", "end").splitlines()
            if question.strip()
        ]

        if not students:
            messagebox.showwarning(
                self.language.get("warning_title"),
                self.language.get("no_students_found")
            )
            return

        if not questions:
            messagebox.showwarning(
                self.language.get("warning_title"),
                self.language.get("no_questions_found")
            )
            return

        questions_per_student = int(self.cmb_questions_per_student.get())

        required_questions = len(students) * questions_per_student

        if len(questions) < required_questions:
            messagebox.showwarning(
                self.language.get("warning_title"),
                self.language.get("not_enough_questions")
            )
            return

        assignment_manager = AssignmentManager()

        assignments = assignment_manager.assign_questions(
            students=students,
            questions=questions,
            questions_per_student=questions_per_student
        )

        excel_generator = ExcelGenerator(self.language)

        filepath = excel_generator.generate(
            students=students,
            assignments=assignments,
            practice_questions=practice_questions,
            questions_per_student=questions_per_student,
            course=self.cmb_course.get(),
            module=self.cmb_module.get(),
            term=self.cmb_term.get(),
            course_language=self.cmb_course_language.get()
        )

        students_count = len(students)
        questions_used = students_count * questions_per_student

        messagebox.showinfo(
            self.language.get("success_title"),
            f"{self.language.get('success_message')}\n\n"
            f"{self.language.get('file')}:\n"
            f"{os.path.basename(filepath)}\n\n"
            f"{self.language.get('location')}:\n"
            f"{filepath}\n\n"
            f"{self.language.get('summary')}\n"
            f"────────────────────\n"
            f"{self.language.get('students')}: {students_count}\n"
            f"{self.language.get('questions_used')}: {questions_used}\n"
            f"{self.language.get('questions_per_student')}: {questions_per_student}"
        )

        print("Excel generated.")

    # =====================================================
    # Update Language
    # =====================================================
    def change_language(self, choice):

        language_code = self.config_data["lists"]["application_languages"][choice]

        self.language.load_language(language_code)

        self.update_ui_language()
    
    def update_ui_language(self):

        self.title(
            f'{self.language.get("title")} v{self.config_data["application"]["version"]}'
        )

        self.lbl_title.configure(text=self.language.get("title"))
        self.lbl_application_language.configure(
            text=self.language.get("application_language")
        )
        self.lbl_course.configure(text=self.language.get("course"))
        self.lbl_module.configure(text=self.language.get("module"))
        self.lbl_term.configure(text=self.language.get("term"))
        self.lbl_year.configure(text=self.language.get("year"))
        self.lbl_course_language.configure(
            text=self.language.get("course_language")
        )
        self.lbl_students.configure(
            text=self.language.get("student_list")
        )
        self.lbl_questions.configure(
            text=self.language.get("question_bank")
        )
        self.lbl_practice_questions.configure(
            text=self.language.get("practice_questions")
        )
        self.btn_generate.configure(
            text=self.language.get("generate_excel")
        )
        self.btn_exit.configure(
            text=self.language.get("exit")
        )

        self.create_menu()

    def show_about(self):
        messagebox.showinfo(
            self.language.get("about"),
            "Question Assignment Tool\n\n"
            "Version 2.1\n\n"
            "Developed by\n"
            "Ing. Lucio M. Buitrón Pareja"
        )
    
    def show_instructions(self):

        window = ctk.CTkToplevel(self)
        window.transient(self)
        window.grab_set()
        window.title(self.language.get("instructions_title"))

        window.geometry("570x300")
        window.update_idletasks()
        window_width = 570
        screen_width = self.winfo_screenwidth()
        x = (screen_width - window_width) // 2
        y = window.winfo_y()
        window.geometry(f"{window_width}x300+{x}+{y}")

        window.resizable(False, False)

        instructions = "\n\n".join([
            self.language.get("instruction_1"),
            self.language.get("instruction_2"),
            self.language.get("instruction_3"),
            self.language.get("instruction_4"),
            self.language.get("instruction_5"),
            self.language.get("instruction_6")
        ])

        label = ctk.CTkLabel(
            window,
            text=instructions,
            justify="left",
            anchor="w"
        )

        label.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=(30, 10)
        )

        btn_close = ctk.CTkButton(
            window,
            text=self.language.get("exit"),
            command=window.destroy
        )

        btn_close.pack(
            pady=(0, 20)
        )