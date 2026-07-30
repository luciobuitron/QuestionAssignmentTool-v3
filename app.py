from ui.main_window import MainWindow


class QuestionAssignmentTool:

    def __init__(self):
        self.window = MainWindow()

    def run(self):
        self.window.mainloop()