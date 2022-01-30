# -*- coding: utf-8 -*-

from odf.opendocument import OpenDocumentText
from odf.style import (Style, TextProperties, ParagraphProperties, ListLevelProperties, TabStop, TabStops)
from odf.text import (H, P, List, ListItem, ListStyle, ListLevelStyleNumber, ListLevelStyleBullet)
from odf import teletype

textdoc = OpenDocumentText()

h1style = Style(name="CenterHeading 1", family="paragraph")
h1style.addElement(ParagraphProperties(textalign="center"))
h1style.addElement(TextProperties(fontsize="18pt", fontweight="bold"))

boldstyle = Style(name="bold", family="text")
boldstyle.addElement(TextProperties(fontweight="bold"))

justifystyle = Style(name="justified", family="paragraph")
justifystyle.addElement(ParagraphProperties(textalign="center"))

numberedliststyle = ListStyle(name="NumberedList")
level = 1
numberedlistproperty = ListLevelStyleNumber(level=str(level), numsuffix=".", startvalue=1)
numberedlistproperty.addElement(ListLevelProperties(minlabelwidth="%fm" % (level)))
numberedliststyle.addElement(numberedlistproperty)

bulletedliststyle = ListStyle(name="BulltList")
level = 1
bulletlistproperty = ListLevelStyleBullet(level=str(level), bulletchar=u".")
bulletlistproperty.addElement(ListLevelProperties(minlabelwidth="%fcm" % level))
bulletedliststyle.addElement(bulletlistproperty)

