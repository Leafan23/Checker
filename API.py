from win32com.client import gencache, Dispatch
from Classes import File, Part, Pdf, Assemble
import os


class API:
    def __init__(self):
        self.api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = Dispatch("KOMPAS.Application.7")
        self.application.Visible = True

        self.main_tree: list[File] = []

        self.documents = self.application.Documents
        self.path = ''
        self.drawing_number  = ''
        self.drawing_name = ''
        self.document = None
        self.part_7 = None

    def open(self, path):

        #Добавить проверку на существующий путь

        document = self.documents.Open(path, 1, 0)
        self.document = document
        self.path = path

        if document.DocumentType == 4 or document.DocumentType == 5:
            kompas_document_3d = self.api7.IKompasDocument3D(document)
            self.part_7 = kompas_document_3d.TopPart
            self.drawing_number = self.get_property_value('Обозначение')
            self.drawing_name = self.get_property_value('Наименование')

    def get_property_value(self, property_name):
        property_mng = self.api7.IPropertyMng(self.application)

        property_keeper = self.api7.IPropertyKeeper(self.part_7)
        for i in range(property_mng.PropertyCount(self.document)):  # ищем свойство по наименованию
            property = property_mng.GetProperty(self.document, i)
            if property.Name == property_name:
                property_value = property_keeper.GetPropertyValue(property, "", True, True)
                return property_value[1]
        return False

    def add_to_main_tree(self, path):
        #self.open(path)
        if os.path.splitext(path)[1] == '.m3d':
            self.main_tree.append(Part(self, id_number=len(self.main_tree)))
        elif os.path.splitext(path)[1] == '.a3d':
            self.main_tree.append(Assemble(self, id_number=len(self.main_tree)))

    def scan(self, path):
        queue_for_check = []
        # открыть сборку
        # добавить в список

        # если есть детали, открыть их поочередно и добавить в список с добавлением родителя изначальной сборки, но только если они небыли раньше обработаны
        # закрыть детали после добавления
        # посчитать количество подсборок (например 3)
        # добавить подсборки в таблицу и занести их адрес в список на проверку
        # закрыть основную сборку
        # открыть сборку из списка и удалить адрес из этого же списка
        # повторять пока список не будет пуст

        self.open(path)
        self.add_to_main_tree(path)



if __name__ == '__main__':
    api = API()
    api.open(r"C:\Users\Leafan\Desktop\Деталь.m3d")
    api.get_property_value('Наименование')
    print(api.get_property_value('Наименование'))

