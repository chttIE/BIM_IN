# coding: utf8

import System
import traceback

from Autodesk.Revit.Exceptions import InvalidOperationException
from pyrevit import UI, forms


class CustomizableEvent(object):
    def __init__(self):
        custom_handler = _CustomHandler()
        custom_handler.customizable_event = self

        self.custom_event = UI.ExternalEvent.Create(custom_handler)

        self.function_or_method = None
        self.args = ()
        self.kwargs = {}

        self.logger = None
        self.last_error = None
        self.last_result = None
        self.is_running = False

    def _raised_method(self):
        if not self.function_or_method:
            return

        self.function_or_method(*self.args, **self.kwargs)

    def raise_event(self, function_or_method, *args, **kwargs):
        if self.is_running:
            self.last_result = "Denied: previous event is still running"
            return self.last_result

        self.function_or_method = function_or_method
        self.args = args
        self.kwargs = kwargs
        self.last_error = None

        try:
            result = self.custom_event.Raise()
            self.last_result = result
            return result

        except Exception as ex:
            self.last_error = traceback.format_exc()
            self.last_result = "Raise failed"

            if self.logger:
                self.logger.exception("Failed to raise customizable event")

            try:
                forms.alert(
                    "Ошибка запуска ExternalEvent:\n\n{}".format(str(ex)),
                    title="CustomizableEvent"
                )
            except:
                pass

            return self.last_result


class _CustomHandler(UI.IExternalEventHandler):
    def __init__(self):
        self.customizable_event = None

    def Execute(self, application):
        ev = self.customizable_event

        if not ev:
            return

        ev.is_running = True

        try:
            ev._raised_method()

        except (InvalidOperationException, Exception, System.Exception) as ex:
            ev.last_error = traceback.format_exc()

            if ev.logger:
                ev.logger.exception("Failed to execute customizable event")

            try:
                forms.alert(
                    "Ошибка выполнения события:\n\n{}".format(str(ex)),
                    title="CustomizableEvent Execute"
                )
            except:
                pass

        finally:
            ev.is_running = False

    def GetName(self):
        return "Execute function or method in IExternalEventHandler"