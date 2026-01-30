# -*- coding: utf-8 -*-
# !usr/bin/env python

import pythoncom, win32com.server.util
from win32com.client import Dispatch, gencache
from tkinter import Tk
from tkinter.ttk import Frame, Button
from connect import SimpleConnection


def on_app_notify_event(command_id, params):
    print("on_app_notify_event call with ", command_id, params)


class MainDialog(Frame):
    def __init__(self, master=None):
        self.__root = Tk()
        self.__root.resizable(0, 0)
        self.__root.title("Тест подписки на события")
        self.__root.wm_attributes("-topmost", 1)
        Frame.__init__(self, master)

        api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        application = Dispatch("KOMPAS.Application.7")
        application.Visible = True
        app_notify_event = BaseEvent(api7.ksKompasObjectNotify, on_app_notify_event, application)
        app_notify_event.advise()


class BaseEvent(object):
    _public_methods_ = ["__on_event"]

    def __init__(self, event, event_handler, event_source):

        self.__event = event
        self.__event_handler = event_handler
        self.__connection = None
        self.event_source = event_source

    def __del__(self):

        if not (self.__connection is None):
            self.__connection.Disconnect()
            del self.__connection

    def _invokeex_(self, command_id, locale_id, flags, params, result, exept_info):
        return self.__on_event(command_id, params)

    def _query_interface_(self, iid):
        if iid == self.__event.CLSID:
            return win32com.server.util.wrap(self)

    def advise(self):
        if self.__connection is None and not (self.event_source is None):
            self.__connection = SimpleConnection(self.event_source, self, self.__event.CLSID) # изменен с win32com.client.connect.SimpleConnection(self.event_source, self, self.__event.CLSID)

    def unadvise(self):
        if self.__connection is not None and self.event_source is not None:
            self.__connection.Disconnect()
            self.__connection = None

    def __on_event(self, command_id, params):
        return self.__event_handler(command_id, params)


if __name__ == "__main__":
    main_app = MainDialog()
    main_app.mainloop()
