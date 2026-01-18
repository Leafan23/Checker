from win32com.client import gencache, Dispatch


class API:
    def __init__(self):
        self.api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = Dispatch("KOMPAS.Application.7")
        self.application.Visible = True
        self.documents = self.application.Documents
        #self.kompas_document_3d = self.api7.IKompasDocument3D(self.document)

    def open(self, path):
        self.document = self.documents.Open(path, 1, 0)

if __name__ == '__main__':
    api = API()
    print('компас открыт')
    api.open(r"C:\Users\Leafan\Desktop\Деталь.m3d")
    print(r'компас открыт файл C:\Users\Leafan\Desktop\Деталь.m3d')
    #api.application.Quit()
    #print('компас закрыт')

    pass