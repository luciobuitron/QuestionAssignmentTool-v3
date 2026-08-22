import random


class AssignmentManager:

    def assign_questions(
        self,
        students,
        questions,
        questions_per_student
    ):
        available_questions = questions.copy()
        random.shuffle(available_questions)

        assignments = {}

        for student in students:

            assigned_questions = []

            for _ in range(questions_per_student):
                assigned_questions.append(
                    available_questions.pop()
                )

            assignments[student] = assigned_questions

        return assignments