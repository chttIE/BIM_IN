# -*- coding: utf-8 -*-"
from models import create_ws_for_links
doc = __revit__.ActiveUIDocument.Document  # noqa
create_ws_for_links(doc)