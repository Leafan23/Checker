class File:
    def __init__(self):
        self.id = 0
        self.path = ''
        self.type = 0


class Pdf(File):
    def __init__(self):
        super().__init__()
        self.type = 1


class Part(File):
    def __init__(self):
        super().__init__()
        self.type = 2
        self.parent = None
        self.no_drawing = False


class Assemble(File):
    def __init__(self):
        super().__init__()
        pass