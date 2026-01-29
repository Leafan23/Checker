from win32com.client import gencache, Dispatch
from Classes import File, Part, Pdf, Assemble, Files
from typing import Any
import os

#TODO добавить обработчик отсутствующих файлов сборки

class API:
    def __init__(self):
        self.api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = Dispatch("KOMPAS.Application.7")
        self.application.Visible = True

        self.files = Files(self)

        self.documents = self.application.Documents
        self.path = ''
        self.drawing_number  = ''
        self.drawing_name = ''
        self.document = None
        self.part_7 = None
        self.model_objects = None

        self.assemble_documents_for_scan: list[Assemble_in_queue] = []

    # принимает путь до файла в виде строки
    def open(self, path: str, parent_id: int=None) -> bool | Any:
        #TODO добавить открытие без проверок и исправлений

        # добавить проверку на поддерживаемые типы файлов
        if not os.path.exists(path):
            return False
        document = self.documents.Open(path, 1, 0)

        self.document = document
        self.path = path

        self.remove_unavailable_documents(document)


        if document.DocumentType == 4 or document.DocumentType == 5:

            kompas_document_3d = self.api7.IKompasDocument3D(document)
            self.part_7 = kompas_document_3d.TopPart
            self.drawing_number = self.get_property_value('Обозначение')
            self.drawing_name = self.get_property_value('Наименование')

            self.files.add_file(path)
            self.files.last_added().parent = parent_id
            parent_for_documents = self.files.last_added().id

            attached_documents = self.check_attached_documents(document)
            product_data_manager = self.api7.IProductDataManager(document)
            property_keeper = self.api7.IPropertyKeeper(self.part_7)

            # добавление в структуру Files
            if document.DocumentType == 4:
                if attached_documents is not None:
                    # проверка привязанных документов
                    for attached_document in attached_documents:
                        if os.path.splitext(path)[0] == os.path.splitext(attached_document)[0]: # path == i без расширения
                            self.files.add_file(attached_document)
                            self.files.last_added().parent = parent_for_documents
                            self.files.id_return(parent_for_documents).add_child(self.files.last_added().id)
                else: # поиск чертежа
                    if self.find_cdw(document):
                        self.files.add_file(self.find_cdw(document))
                        self.files.last_added().parent = parent_for_documents
                        self.files.id_return(parent_for_documents).add_child(self.files.last_added().id)
                        product_data_manager.SetObjectAttachedDocuments(property_keeper, self.find_cdw(document))

            if document.DocumentType == 5:
                documents_for_attach = []
                if attached_documents is not None:
                    for i in attached_documents:
                        documents_for_attach.append(i)
                spw = self.find_spw(document)
                cdw = self.find_cdw(document)

                if spw:
                    self.files.add_file(spw)
                    self.files.last_added().parent = parent_for_documents
                    self.files.id_return(parent_for_documents).add_child(self.files.last_added().id)
                    documents_for_attach.append(spw)
                if cdw:
                    self.files.add_file(cdw)
                    self.files.last_added().parent = parent_for_documents
                    self.files.id_return(parent_for_documents).add_child(self.files.last_added().id)
                    documents_for_attach.append(cdw)

                product_data_manager.SetObjectAttachedDocuments(property_keeper, documents_for_attach)

            #TODO при открытии чертежа или спецификации проверить соответствие обозначению в файле и в основной надписи
            #TODO при открытии чертежа, проверить наличие спецификации на чертеже, проверить обозначение на соответствие имени файла
            #TODO при открытии спецификации, проверить на соответствие обозначения

            document.Save()
            return parent_for_documents
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

    def scan(self, path=None): #TODO сделать обработку исполнений, пока работает только файлами
        global parent
        queue_for_check_parts = []
        queue_for_check_assembles = []

        # перебор отсканированного списка
        if path is None:
            for i in self.assemble_documents_for_scan:
                if i.count_of_path == 0: continue
                path = i.next_path()
                if path != None: parent = self.open(path, i.id_of_master)
                break
            else:
                return
        else:
            parent = self.open(path)

        # получение списков для перебора
        feature_7 = self.api7.IFeature7(self.part_7)
        for i in feature_7.SubFeatures(0, True, False):
            if i.ModelObjectType == 104:
                part_7 = self.api7.IPart7(i)
                if part_7.Detail:
                    queue_for_check_parts.append(part_7.FileName)
                else:
                    queue_for_check_assembles.append(part_7.FileName)

        # удаление из списка повторяющихся файлов
        queue_for_check_parts = list(set(queue_for_check_parts))
        queue_for_check_assembles = list(set(queue_for_check_assembles))

        # добавление файла в список
        self.assemble_documents_for_scan.append(Assemble_in_queue(parent, queue_for_check_assembles))

        # перебор всех сборок и деталей
        self.document.Close(1)
        for i in queue_for_check_parts:
            self.open(i, parent)
            self.document.Close(1)
        self.scan()


    # костыль для удаления недействительных документов
    def remove_unavailable_documents(self, document): #TODO переработать функцию на более оптимальную
        if document.DocumentType == 4 or document.DocumentType == 5:
            available_documents = []
            attached_documents = self.check_attached_documents(document)
            kompas_document_3d = self.api7.IKompasDocument3D(document)
            part_7 = kompas_document_3d.TopPart
            if attached_documents:
                for path_to_document in attached_documents:
                    if os.path.exists(path_to_document):
                        available_documents.append(path_to_document)

                # удалить все документы
                product_data_manager = self.api7.IProductDataManager(document)
                product_objects = []
                if isinstance(product_data_manager.ProductObjects(1), tuple):
                    product_objects += product_data_manager.ProductObjects(1)
                else:
                    product_objects.append(product_data_manager.ProductObjects(1))
                for product in product_objects:
                    if 'IPropertyKeeper' in str(product):
                        unique_meta_object_key = product.UniqueMetaObjectKey
                        product_data_manager.DeleteProductObject(unique_meta_object_key)

                # добавить вернуть документы в список
                product_data_manager.SetObjectAttachedDocuments(part_7, available_documents)
            return True
        else:
            return False


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

    def find_pdf(self): #TODO написать функцию
        pass

    def find_spw(self, document):
        if document.DocumentType == 5:
            if os.path.exists(document.Path + os.path.splitext(document.Name)[0] + '.spw'):
                return document.Path + os.path.splitext(document.Name)[0] + '.spw'
        return None

    def find_cdw(self, document) -> str | None:
        """
        Функция ищет чертеж с форматом cdw.
        Принимает: IKompasDocument
        Возвращает: путь до файла или None
        """
        if document.DocumentType == 4:
            if os.path.exists(document.Path + os.path.splitext(document.Name)[0] + '.cdw'):
                return document.Path + os.path.splitext(document.Name)[0] + '.cdw'
            return None
        elif document.DocumentType == 5:
            if os.path.exists(document.Path + os.path.splitext(document.Name)[0] + ' СБ.cdw'):
                return document.Path + os.path.splitext(document.Name)[0] + ' СБ.cdw'
            #TODO Дописать возврат нескольких значений и исправить обработку функций в программе
            return None
        else:
            return None

class Assemble_in_queue:
    def __init__(self, id: int, child_path: list[str]):
        self.child_path = child_path
        self.id_of_master = id
        self.count_of_path = len(self.child_path)
        self.list_fo_return = self.child_path

    def next_path(self) -> str | None:
        if len(self.list_fo_return) > 0:
            self.count_of_path -= 1
            return self.list_fo_return.pop(0)
        else:
            return None


if __name__ == '__main__':
    api = API()
    api.open(r"C:\Users\Leafan\Desktop\Деталь.m3d")
    api.get_property_value('Наименование')
    print(api.get_property_value('Наименование'))

