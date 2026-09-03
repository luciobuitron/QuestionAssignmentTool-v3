# Software Design Document (SDD)

# Question Assignment Tool (QAT)

**Version:** 3.0  
**Status:** Stable  
**Author:** Ing. Lucio M. Buitrón Pareja

## 1. Project Overview

The Question Assignment Tool (QAT) is a desktop application designed to simplify the assignment and evaluation of theoretical and practical questions for students.

The application allows instructors to:

- Manage student groups and student lists.
- Manage a bank of theoretical questions.
- Automatically assign theoretical questions without repeating a question for the same student.
- Optionally manage and assign practice questions by student group.
- Use question tags to improve theoretical question distribution when tags are provided.
- Configure the number of theoretical questions assigned to each student.
- Configure the Total Grade and its Defense/Theory percentage distribution.
- Generate a professionally formatted Microsoft Excel workbook for student evaluation.
- Evaluate students using predefined grade scales and dropdown lists.
- Calculate Defense, Theory, Practice, and Final Total Grade values automatically.
- Protect Excel worksheets while allowing evaluators to modify only the required fields.
- Support English, Spanish, and Portuguese user interfaces.
- Record application events and errors in a persistent log file.

The application is designed to run completely offline and can be packaged as a standalone Windows executable.


## 2. Objectives

The main objectives of the Question Assignment Tool (QAT) are:

- Simplify the process of assigning theoretical questions to students.
- Provide a controlled and consistent method for distributing questions among student groups.
- Prevent the same theoretical question from being assigned more than once to the same student.
- Support question tags to improve assignment diversity when sufficient tagged questions are available.
- Allow instructors to configure the number of theoretical questions assigned to each student.
- Support optional practice questions assigned at the student-group level.
- Provide a configurable Total Grade with independent Defense and Theory percentage distribution.
- Automate grade calculations for Defense, Theory, Practice, and Final Total Grade.
- Generate a structured and professionally formatted Excel workbook for student evaluation.
- Restrict editing of the generated workbook to the fields that require evaluator input.
- Provide controlled dropdown lists and data validation to reduce invalid evaluation entries.
- Support English, Spanish, and Portuguese interfaces.
- Provide persistent application logging for successful operations, validations, warnings, and errors.
- Reduce manual work and improve consistency in the student evaluation process.


## 3. Functional Requirements

The Question Assignment Tool (QAT) shall provide the following functional capabilities:

### 3.1 Student Management

- Allow the instructor to enter a list of students.
- Support multiple student groups.
- Use blank lines to separate student groups.
- Preserve the group structure when generating the Excel workbook.
- Support groups with different numbers of students.

### 3.2 Theoretical Question Management

- Allow the instructor to enter a list of theoretical questions.
- Support optional tags associated with theoretical questions.
- Allow the instructor to configure the number of theoretical questions assigned to each student.
- Ensure that the same theoretical question is not assigned more than once to the same student.
- Use question tags to improve assignment diversity when tags are available.
- If there are not enough distinct tagged questions, allow tag repetition between students while maintaining the restriction that a student does not receive the same question more than once.
- Reject configurations containing mixed tagged and untagged theoretical questions.
- Reject invalid tag configurations.

### 3.3 Practice Question Management

- Allow the instructor to optionally enter a list of practice questions.
- Assign one practice question per student group.
- Allow the instructor to select the practice question for each group directly from the generated Excel workbook.
- Allow the evaluator to assign a practice grade to each student.
- Support the following practice grades:
  - `RESOLVED`
  - `PARTIAL`
  - `NOT RESOLVED`
- Automatically calculate the practice grade value according to the selected practice grade.

### 3.4 Grade Configuration

- Allow the instructor to configure the Total Grade.
- Allow the instructor to define the percentage distribution between Defense and Theory.
- Validate that the Defense and Theory percentages form a valid distribution.
- Support predefined theoretical and defense grade levels:
  - `EXCELLENT`
  - `GOOD`
  - `FAIR`
  - `POOR`
  - `NOT GRADED`
- Automatically calculate the corresponding Defense and Theory grade values.
- Calculate the Final Total Grade from Defense, Theory, and Practice values.

### 3.5 Excel Workbook Generation

- Generate a Microsoft Excel workbook containing the assigned questions and evaluation fields.
- Generate a `Theoretical Questions` worksheet.
- Generate a `Practice Questions` worksheet when practice questions are provided.
- Generate supporting hidden worksheets used for dropdown lists and data validation.
- Apply formulas for automatic grade calculations.
- Apply formatting, borders, merged cells, filters, and frozen rows as required.
- Apply alternating colors to student and group sections to improve readability.
- Generate a timestamped filename based on the course, module, term, course language, and generation timestamp.

### 3.6 Data Validation and Protection

- Provide dropdown lists for evaluator-controlled fields.
- Reject values that are not included in the corresponding dropdown lists.
- Use Excel `STOP` error validation for invalid dropdown values.
- Protect generated worksheets against unintended modification.
- Keep calculated and informational cells locked.
- Keep evaluator input fields editable.
- Allow manual entry of Practice Question Grade on the theoretical worksheet when no practice questions are configured.
- Allow comments to be entered in the designated comment fields.


### 3.7 Logging

- Record successful Excel generation events.
- Record validation events.
- Record warnings.
- Record application errors.
- Store log entries persistently in the application log file.
- Include detailed error information and traceback data for unexpected generation errors.


## 4. Non-Functional Requirements

### 4.1 Usability

- Provide a simple and clear graphical user interface.
- Minimize manual work during question assignment and evaluation.
- Provide clear validation and error messages.

### 4.2 Reliability

- Prevent invalid configurations from generating incomplete or inconsistent workbooks.
- Preserve assignment rules and grade calculations consistently.
- Record unexpected errors in the application log.

### 4.3 Data Integrity

- Prevent unintended modification of calculated and informational Excel cells.
- Validate evaluator input through controlled dropdown lists.
- Maintain the relationship between assigned questions, students, groups, and calculated grades.

### 4.4 Compatibility

- Run on Windows environments.
- Generate standard Microsoft Excel `.xlsx` files.
- Support English, Spanish, and Portuguese.

### 4.5 Maintainability

- Use a modular architecture separating the user interface, assignment logic, logging, and Excel generation.
- Keep configuration and language resources separate from application logic.
- Allow future enhancements without affecting the core assignment process.


## 5. User Interface

The QAT user interface provides the main configuration and generation workflow.

### 5.1 Main Configuration

The interface allows the instructor to configure:

- Course, module, term, year, and course language.
- Student groups and student lists.
- Theoretical questions.
- Questions per student.
- Total Grade.
- Defense/Theory percentage distribution.
- Optional practice questions.

### 5.2 Actions

The main interface provides controls to:

- Generate the Excel assignment workbook.
- Change the application language.
- View application information.
- View usage instructions.

### 5.3 Supported Languages

The interface supports:

- English
- Spanish
- Portuguese

UI labels, instructions, and messages are updated according to the selected language.


## 6. Validation Rules

QAT validates the input configuration before generating the Excel workbook.

### 6.1 Student Validation

- At least one student must be provided.
- Blank lines are used only to separate student groups.

### 6.2 Theoretical Question Validation

- At least one theoretical question must be provided.
- The number of available questions must be sufficient for the configured assignment.
- Theoretical questions must use either valid tags consistently or no tags.
- Mixed tagged and untagged questions are not allowed.
- Invalid tag configurations are rejected.

### 6.3 Grade Validation

- Total Grade must be greater than zero.
- Defense and Theory percentages must form a valid distribution.
- The configured values must allow the requested assignment to be generated.

### 6.4 Practice Validation

- Practice questions are optional.
- When provided, at least one practice question must be available.
- Practice question grades must use the predefined valid values.

If a validation fails, QAT displays an appropriate warning and prevents Excel generation.


## 7. Assignment Rules

QAT assigns theoretical questions according to the configured number of questions per student.

- Questions are assigned automatically to each student.
- The same theoretical question is never assigned more than once to the same student.
- When valid tags are provided, QAT attempts to assign different tags to the same student.
- If there are not enough distinct tagged questions, tag repetition between students is allowed.
- A question may be assigned to different students when required by the available question pool.
- The assignment process preserves the original student group structure.
- Practice questions are assigned at the group level, with one selected practice question per group.


## 8. Excel Specification

QAT generates a formatted `.xlsx` workbook for student evaluation.

### 8.1 Theoretical Questions Worksheet

The `Theoretical Questions` worksheet contains:

- Group
- Student number and name
- Defense Grade and calculated value
- Theoretical Question
- Theoretical Question Grade and calculated value
- Practice Question Grade
- Final Total Grade
- Comments

The worksheet also includes:

- Grade reference tables.
- Automatic formulas for grade calculations.
- Dropdown lists and data validation.
- Protected cells and editable evaluator fields.
- Frozen rows, filters, merged cells, borders, and alternating colors.

### 8.2 Practice Questions Worksheet

The `Practice Questions` worksheet is generated only when practice questions are provided.

It contains:

- Group
- Student number and name
- Practice Question
- Practice Question Grade
- Practice Grade Value
- Comments

Practice grades are calculated as:

- `RESOLVED` → 10% of Total Grade
- `PARTIAL` → 5% of Total Grade
- `NOT RESOLVED` → 0%

### 8.3 Grade Calculation

The Final Total Grade is calculated from:

`Defense Grade Value + Theoretical Grade Value + Practice Grade Value`

When no practice questions are configured, the Practice Grade Value is treated as zero.


## 9. File Naming

Generated Excel files use the following naming convention:

`<Course>_<Module>_Term<Term>_<CourseLanguage>_<YYYY-MM-DD_HH-MM-SS>.xlsx`

Spaces are removed from the course and module names.

Example:

`OperatingSystems2_Module1_Term1_Spanish_2026-09-02_23-14-41.xlsx`

The timestamp ensures that each generated workbook has a unique filename.


## 10. Internationalization

QAT supports the following interface languages:

- English
- Spanish
- Portuguese

Language resources are stored separately from the application logic.

Changing the selected language updates the main interface, instructions, validation messages, and other user-facing text accordingly.


## 11. Configuration

QAT uses configuration and language resources to define application settings and user-facing text.

The configuration supports:

- Default application settings.
- Interface language.
- Application preferences used during execution.

Assignment-specific values such as the number of questions per student, Total Grade, and Defense/Theory distribution are configured directly through the user interface.


## 12. Architecture

QAT follows a modular architecture that separates the main application responsibilities.

The main components are:

- **Main Window:** manages the graphical user interface and user interaction.
- **Assignment Manager:** handles theoretical question assignment and tag-based distribution.
- **Excel Generator:** creates and formats the Excel workbook, including formulas, validations, and protection.
- **Logger:** records application events, validations, warnings, and errors.
- **Language Resources:** provide multilingual interface text.

The general workflow is:

`Main Window → Input Validation → Assignment Manager → Excel Generator → Excel Workbook`

Logging is used throughout the application to record relevant events and errors.


## 13. Logging

QAT maintains a persistent application log at:

`logs/qat.log`

The logger records:

- Successful operations.
- Validation events.
- Warnings.
- Errors.

Each log entry includes a timestamp, event level, and message.

Unexpected errors also include traceback information to facilitate troubleshooting.


## 14. Data Protection

Generated Excel worksheets are protected to prevent unintended modifications.

### Editable Fields

The evaluator can modify only the fields required during evaluation:

**Theoretical Questions:**
- Group
- Defense Grade
- Theoretical Question Grade
- Practice Question Grade when no Practice Questions worksheet is generated
- Comments

**Practice Questions:**
- Group
- Practice Question
- Practice Question Grade
- Comments

### Protected Fields

Calculated values, assigned questions, student information, formulas, reference tables, and other supporting data are protected against direct modification.

Dropdown fields use Excel Data Validation with `STOP` error handling to prevent invalid values.


## 15. Version History

| Version | Status | Description |
|---|---|---|
| 1.0 | Stable | Initial version of QAT with basic theoretical question assignment and Excel generation. |
| 2.1 | Stable | Base functional version used as the foundation for QAT V3. |
| 3.0 | Stable | Added configurable questions per student, tag-based assignment, optional practice questions, Total Grade calculation, Defense/Theory distribution, Excel protection, data validation, multilingual support, logging, and enhanced evaluation features. |


## 16. Future Enhancements

The following improvements are planned for future versions of QAT:

- Support a three-component grade distribution between Defense, Theory, and Practice.
- Allow Practice Questions to be assigned individually to each student instead of by group.
- Make the grade scale configurable through the configuration file.
- Evaluate a numeric/read-only design for Defense and Theory grade fields.
- Refactor and simplify the Excel generation code.
- Refactor the main window generation workflow to improve separation of responsibilities.
- Investigate the intermittent Excel dropdown rendering issue that can cause inconsistent dropdown widths.