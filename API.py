from win32com.client import gencache, Dispatch
from Classes import File, Part, Pdf, Assemble
import os

from test import check_attached_documents


class API:
    def __init__(self):
        self.api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = Dispatch("KOMPAS.Application.7")
        self.application.Visible = False

        self.main_tree: list[File] = []

        self.documents = self.application.Documents
        self.path = ''
        self.drawing_number  = ''
        self.drawing_name = ''
        self.document = None
        self.part_7 = None
        self.model_objects = None

    # принимает путь до файла в виде строки
    def open(self, path):

        # добавить проверку на поддерживаемые типы файлов
        attached_documents = []
        if not os.path.exists(path):
            return False

        document = self.documents.Open(path, 1, 0)
        self.add_to_main_tree(path)
        self.document = document
        self.path = path

        if document.DocumentType == 4 or document.DocumentType == 5:
            kompas_document_3d = self.api7.IKompasDocument3D(document)
            self.part_7 = kompas_document_3d.TopPart
            self.drawing_number = self.get_property_value('Обозначение')
            self.drawing_name = self.get_property_value('Наименование')

            if document.DocumentType == 4:
                attached_documents = self.check_attached_documents(document)
                if attached_documents is not None:

                    # проверка привязанных документов
                    for i in attached_documents:

                        # проверка на существование документа
                        if os.path.exists(i):
                            print('это: ', os.path.splitext(path)[0], 'равно этому:', os.path.splitext(i)[0])
                            if os.path.splitext(path)[0] == os.path.splitext(i)[0]: # path == i без расширения
                                self.main_tree[-1].drawing = True
                        else:
                            print('Типо удален из документов', i)
                            pass # тут надо удалить документ из привязанных
                else:
                    if self.find_cdw(document):
                        product_data_manager = self.api7.IProductDataManager(document)
                        property_keeper = self.api7.IPropertyKeeper(self.part_7)
                        product_data_manager.SetObjectAttachedDocuments(property_keeper, self.find_cdw(document))
                        document.Save() #TODO Убедится, что нужно сохранение. так как в .Close это предусмотрено
        document.Close(1)

                        # проверка на
        return True

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
        queue_for_check_parts = []
        queue_for_check_assembles = []
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

        feature_7 = self.api7.IFeature7(self.part_7)

        for i in feature_7.SubFeatures(0, True, False):
            if i.ModelObjectType == 104:
                print(i.ModelObjectType)
                print(i.Name)
                part_7 = self.api7.IPart7(i)
                print(part_7.FileName)
                print('Это деталь? ', part_7.Detail)
                print('Обозначение: ', part_7.Marking)
                print('Количество вставок: ', part_7.InstanceCount)
                print('-----------------------')
                if part_7.Detail:
                    queue_for_check_parts.append(part_7.FileName)
                else:
                    queue_for_check_assembles.append(part_7.FileName)

        # удаление из списка повторяющихся файлов
        queue_for_check_parts = list(set(queue_for_check_parts))
        queue_for_check_assembles = list(set(queue_for_check_assembles))
        print(queue_for_check_parts)
        print(queue_for_check_assembles)

        for i in queue_for_check_parts:
            self.open(i)

    # Принимает на вход класс IKompasDocument.
    # Обрабатывает файл, и выдает все привязанные файлы в виде списка
    def check_attached_documents(self, document):
        documents_array = () #TODO проверить на ненужность documents_array = ()
        document_3d = self.api7.IKompasDocument3D(document)
        product_data_manager = self.api7.IProductDataManager(document)

        # возвращает спецификации, только из 3Д документов
        if document.DocumentType == 4 or document.DocumentType == 5:
            part_7 = self.api7.IPart7(document_3d.TopPart)
            property_keeper = self.api7.IPropertyKeeper(part_7)
            if product_data_manager.ObjectAttachedDocuments(property_keeper) is not None:
                documents_array = product_data_manager.ObjectAttachedDocuments(property_keeper) + documents_array

        # возвращает все документы включая ссылку на себя же
        property_keeper = self.api7.IPropertyKeeper(document)
        documents_array = product_data_manager.ObjectAttachedDocuments(property_keeper) + documents_array

        # преобразовываем в список и удаляем текущий документ из списка
        documents_array = list(documents_array)
        documents_array.remove(document.PathName)
        if not documents_array:
            return None
        return documents_array

    def find_pdf(self):
        pass

    def find_spw(self):
        pass

    def find_cdw(self, document):
        if document.DocumentType == 4:
            if os.path.exists(document.Path + os.path.splitext(document.Name)[0] + '.cdw'):
                return document.Path + os.path.splitext(document.Name)[0] + '.cdw'
            return None
        elif(document.DocumentType == 5):
            pass
        else:
            return None




if __name__ == '__main__':
    api = API()
    api.open(r"C:\Users\Leafan\Desktop\Деталь.m3d")
    api.get_property_value('Наименование')
    print(api.get_property_value('Наименование'))

