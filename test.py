from API import API

def check_attached_documents(document):
    print(document)
    document_3d = kompas.api7.IKompasDocument3D(document)

    part_7 = kompas.api7.IPart7(document_3d.TopPart)
    product_data_manager = kompas.api7.IProductDataManager(document)

    if document.DocumentType == 4 or document.DocumentType == 5:
        iIPropertyKeeper_topPart = kompas.api7.IPropertyKeeper(part_7)
        print(f'Связанные спецификации:\n' + '\n'.join(
            product_data_manager.ObjectAttachedDocuments(iIPropertyKeeper_topPart)) + '\n')

    iIPropertyKeeper_doc = kompas.api7.IPropertyKeeper(document)
    print(f'Список подключенных к объекту документов без спецификаций:\n' + '\n'.join(
        product_data_manager.ObjectAttachedDocuments(iIPropertyKeeper_doc)) + '\n')

if __name__ == '__main__':
    kompas = API()
    kompas_document = kompas.application.ActiveDocument
    check_attached_documents(kompas_document)

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
