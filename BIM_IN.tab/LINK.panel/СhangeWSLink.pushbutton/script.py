# -*- coding: utf-8 -*-"

__title__ = 'Создание\nРН связей' 
__doc__ = """Скрипт смотрит какие связи загружены
проверяет в каких рабочих наборах 
они находятся и если
рн нет, создаст его, поместит в него связь и закрепит её
Так же если связь имеет в имени ZONES - уберет у нее галочку видимости"""
___author__ = 'Leonid Malyshev\
IliaNistratov'
__highlight__ = 'updated'
from logIN import lg
from models import create_ws_for_links


doc = __revit__.ActiveUIDocument.Document  # noqa


lg(doc,__title__)   
create_ws_for_links(doc)