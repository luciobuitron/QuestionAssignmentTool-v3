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
        all_tags_different = True

        for student in students:

            assigned_questions = []
            assigned_tags = set()

            for _ in range(questions_per_student):

                selected_index = None

                # Try to find a question with a different tag
                for index, question in enumerate(available_questions):

                    tag_end = question.find("]")

                    if tag_end > 1:
                        tag = question[1:tag_end].strip()

                        if tag not in assigned_tags:
                            selected_index = index
                            break

                # If no different tag is available,
                # use any remaining question
                if selected_index is None:
                    selected_index = 0
                    all_tags_different = False

                question = available_questions.pop(selected_index)
                assigned_questions.append(question)

                tag_end = question.find("]")

                if tag_end > 1:
                    tag = question[1:tag_end].strip()
                    assigned_tags.add(tag)

            assignments[student] = assigned_questions

        return assignments, all_tags_different