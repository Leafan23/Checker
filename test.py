#TODO добавить функцию проверки на необходимые файлы, чертеж, спецификация, pdf
import API


# Принимает на вход класс IKompasDocument.
# Обрабатывает файл, и выдает все привязанные файлы в виде списка
def check_attached_documents(document):
    documents_array = ()
    document_3d = kompas.api7.IKompasDocument3D(document)
    product_data_manager = kompas.api7.IProductDataManager(document)

    # возвращает спецификации, только из 3Д документов
    if document.DocumentType == 4 or document.DocumentType == 5:
        part_7 = kompas.api7.IPart7(document_3d.TopPart)
        property_keeper = kompas.api7.IPropertyKeeper(part_7)
        if product_data_manager.ObjectAttachedDocuments(property_keeper) is not None:
            documents_array = product_data_manager.ObjectAttachedDocuments(property_keeper) + documents_array

    # возвращает все документы включая ссылку на себя же
    property_keeper = kompas.api7.IPropertyKeeper(document)
    documents_array = product_data_manager.ObjectAttachedDocuments(property_keeper) + documents_array

    # преобразовываем в список и удаляем текущий документ из списка
    documents_array = list(documents_array)
    documents_array.remove(document.PathName)
    if not documents_array:
        return None
    return documents_array

if __name__ == '__main__':
    kompas = API.API()
    kompas_document = kompas.application.ActiveDocument
    print(check_attached_documents(kompas_document))

'''if __name__ == '__main__':
    kompas = API()

    kompas_document = kompas.application.ActiveDocument
    kompas_document_3d = kompas.api7.IKompasDocument3D(kompas_document)
    part_7 = kompas_document_3d.TopPart
    parts_7 = part_7.Parts

    feature_7 = kompas.api7.IFeature7(part_7)

    property_keeper = kompas.api7.IPropertyKeeper(kompas_document)
    product_data_manager = kompas.api7.IProductDataManager(kompas_document)
    array_of_documents = product_data_manager.ObjectAttachedDocuments(property_keeper)
    for i in array_of_documents:
        print(i)


    for i in feature_7.SubFeatures(0, True, False): # перебор деталей
        if i.ModelObjectType == 104:
            part_7 = kompas.api7.IPart7(i)
            property_mng = kompas.api7.IPropertyMng(kompas.application)

            property_keeper = kompas.api7.IPropertyKeeper(part_7)

            for k in range(property_mng.PropertyCount(kompas_document)):  # перебор свойств
                property = property_mng.GetProperty(kompas_document, k)
                #print(k, property.Name, property_keeper.GetPropertyValue(property, False, False)[1])'''

'''from API import API


def attach_document():
    product_data_manager.SetObjectAttachedDocuments(property_keeper, (r"D:\Projects\Checker\Kompas_files\ГКЮШ.000.000 - Сборка.spw",
                                                                           r"D:\Projects\Checker\Kompas_files\ГКЮШ.000.000 - Сборка.cdw",
                                                                           r"D:\Projects\Checker\Kompas_files\ГКЮШ.000.000 - Сборка — копия.spw",
                                                                           r"D:\Projects\Checker\Kompas_files\ГКЮШ.000.000 - Сборка — копия.cdw"))

kompas = API()
document = kompas.application.ActiveDocument
document_3d = kompas.api7.IKompasDocument3D(document)
part_7 = document_3d.TopPart
property_keeper = kompas.api7.IPropertyKeeper(part_7)
product_data_manager = kompas.api7.IProductDataManager(document)

property_manager = kompas.api7.IPropertyMng(kompas.application)

#attach_document()


for i in range(property_manager.PropertyCount(document)):
    property = property_manager.GetProperty(document, i)
    #print(property.Name, ':    ', property_keeper.GetPropertyValue(property, "", True, True)[1])
for i in product_data_manager.ObjectAttachedDocuments(property_keeper):
    print(i)'''
