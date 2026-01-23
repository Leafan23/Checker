# -*- coding: utf-8 -*-
import pythoncom
from win32com.client import Dispatch, gencache, VARIANT

#  Получи константы
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Получи API интерфейсов версии 5
KAPI = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
iikompas_object = KAPI.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(KAPI.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Получи API интерфейсов версии 7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
kompas_api_object = kompas_api7_module.IKompasAPIObject(
    Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IKompasAPIObject.CLSID,
                                                             pythoncom.IID_IDispatch))

def draw_del():
    iProductDataManager = kompas_api7_module.IProductDataManager(iKompasDocument)
    product_objects = iProductDataManager.ProductObjects(kompas6_constants.ksPOTDocumentObject)
    print(product_objects)
    # try:
        # if 'IPropertyKeeper' in product_objects:
        #     print(product_objects)
        # for product in product_objects:
        #     if 'IPropertyKeeper' in str(product):
        #         unique_meta_object_key = product.UniqueMetaObjectKey
        #         iProductDataManager.DeleteProductObject(unique_meta_object_key)
        # print("Удалил чертежи нормальным способом")
    # except:
    #     print("Plan b")



application = kompas_api_object.Application

#  Получим активный документ
iKompasDocument = application.ActiveDocument
IKompasDocument1 = kompas_api7_module.IKompasDocument1(iKompasDocument)
iKompasDocument3D = kompas_api7_module.IKompasDocument3D(iKompasDocument)
iDocument3D = iikompas_object.ActiveDocument3D()

# Имя файла
iPart7 = iKompasDocument3D.TopPart

#Проверка на тип файла
if iPart7.Detail is True:
    # Получаем список прикрепленных файлов
    GetExternalFilesNamesEx = IKompasDocument1.GetExternalFilesNamesEx(False)
    print(GetExternalFilesNamesEx)
    if GetExternalFilesNamesEx[1] is not None:
        if 31 in GetExternalFilesNamesEx[2]:
            print("Надо удалить ")
            draw_del()
        else:
            print('чертежей нет')
    else:
        print("Вообще ничего нет.")