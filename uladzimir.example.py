# Create a PDF document using itext7 and Jpype
# Minimal Example revised with Uladzimir's comments/code from
# https://stackoverflow.com/a/72385258/9525730
# Attempting to fit everything on one A4 page

import calendar
from datetime import datetime

def createDoc(mon="Mar", year="22"):
    '''Creates a PDF'''

    def _headerCell(text="", r=1, c=1):
        return Cell(r, c).add(Paragraph(str(text)))
        
    def _cell(text="", r=1, c=1):
        return Cell(r, c).add(Paragraph(str(text)))

    # Get days in month
    now = datetime.strptime(mon+"/"+year, "%b/%y")
    days_in_month = calendar.monthrange(now.year, now.month)[1]
    # Create a table with 10 columns
    table = Table(10)
    table.useAllAvailableWidth()
    table.setTextAlignment(TextAlignment.CENTER)
    # ROW 1 HEADER
    table.addHeaderCell(_headerCell("Col 1", 1, 3))  
    table.addHeaderCell(_headerCell("Col 2"))
    table.addHeaderCell(_headerCell("Col 3", 1, 2))
    table.addHeaderCell(_headerCell("Col 4"))
    table.addHeaderCell(_headerCell("Col 5"))
    table.addHeaderCell(_headerCell("Col 6", 1, 2))
    # ROW 2 HEADER
    table.addHeaderCell(_headerCell("Col 1"))
    table.addHeaderCell(_headerCell("Col 2", 1, 2))
    table.addHeaderCell(_headerCell("Col 3"))
    table.addHeaderCell(_headerCell("Col 4", 1, 2))
    table.addHeaderCell(_headerCell("Col 5"))
    table.addHeaderCell(_headerCell("Col 6"))
    table.addHeaderCell(_headerCell("Col 7", 1, 2))
    # ROW 3 HEADER
    table.addHeaderCell(_headerCell())
    table.addHeaderCell(_headerCell("A"))
    table.addHeaderCell(_headerCell("B"))
    table.addHeaderCell(_headerCell("C"))
    table.addHeaderCell(_headerCell("D"))
    table.addHeaderCell(_headerCell("E"))
    table.addHeaderCell(_headerCell("F"))
    table.addHeaderCell(_headerCell("G"))
    table.addHeaderCell(_headerCell("H", 1, 2))
    # ROW 4 HEADER
    table.addHeaderCell(_headerCell("AA", 1, 2))
    table.addHeaderCell(_headerCell("BB"))
    table.addHeaderCell(_headerCell("CC"))
    table.addHeaderCell(_headerCell("DD", 1, 4))
    table.addHeaderCell(_headerCell("EE"))
    table.addHeaderCell(_headerCell("FF"))
    for r in range(days_in_month):
        table.addCell(_cell(r + 1))
        table.addCell(_cell("I"))
        table.addCell(_cell("II"))
        table.addCell(_cell("III"))
        table.addCell(_cell("IV"))
        table.addCell(_cell("V"))
        table.addCell(_cell("VI"))
        table.addCell(_cell("VII"))
        table.addCell(_cell("VIII"))
        table.addCell(_cell("IX"))
        
    table.addCell(_cell("a", 1, 4))
    table.addCell(_cell("b"))
    table.addCell(_cell("c"))
    table.addCell(_cell("d"))
    table.addCell(_cell("e"))
    table.addCell(_cell("f"))
    table.addCell(_cell("g"))
    table.addCell(_cell("h", 1, 4))
    table.addCell(_cell("i"))
    table.addCell(_cell("j"))
    table.addCell(_cell("k"))
    table.addCell(_cell("l"))
    table.addCell(_cell("m"))
    table.addCell(_cell("n"))
    table.addCell(_cell("o", 1, 3))
    table.addCell(_cell("p"))
    table.addCell(_cell("q"))
    table.addCell(_cell("r", 1, 4))
    table.addCell(_cell("s"))
    table.addCell(_cell("t", 1, 3))
    table.addCell(_cell("u"))
    table.addCell(_cell("v"))
    table.addCell(_cell("w", 1, 4))
    table.addCell(_cell("x"))
    table.addCell(_cell("y", 1, 3))
    table.addCell(_cell("z", 1, 4))
    table.addCell(_cell("aa", 1, 3))

    pdfDoc = PdfDocument(PdfWriter(f'{PDF_NAME}.pdf'))
    document = Document(pdfDoc)

    # Source/comments from @Uladzimir via https://stackoverflow.com/a/72385258/9525730
    # Since it's too long to be fully fit via usual layout flow (e.g. Document#add), we should somehow scale it. 
    # But first of all, let us find how much space this table occupies if the page to be placed upon is boundless:
    result = table.createRendererSubTree().setParent(document.getRenderer()).layout(LayoutContext(LayoutArea(1, Rectangle(10000, 10000))))
    occupiedRectangle = result.getOccupiedArea().getBBox()
    # Now let's create a form xobject of this table, which we will scale a few lines below:
    xObject = PdfFormXObject(Rectangle(occupiedRectangle.getWidth(), occupiedRectangle.getHeight()))
    Canvas(xObject, pdfDoc).add(table).close()
    # So now we have the xObject of the table, the only question is how to fit it, e.g. which scale coefficients to apply:
    coefficient = min(PageSize.A4.getWidth() / occupiedRectangle.getWidth(), PageSize.A4.getHeight() / occupiedRectangle.getHeight())
    # We're almost done: now let's add the scaled version of the table to the document's page:
    PdfCanvas(pdfDoc.addNewPage()).saveState().concatMatrix(coefficient, 0, 0, coefficient, 0, 0).addXObject(xObject).restoreState()

    document.close()
    print("Wrote", f"{PDF_NAME}.pdf")


PDF_NAME = "example"

if __name__ == '__main__':
    import jpype
    import jpype.imports
    from jpype.types import *
    print("Starting JVM")
    jpype.startJVM("-Xms256m", "-Xmx256m",
                   "-Doracle.metrics.clientid=JPYPE", classpath=["itext.7.2.2.jar"])

    from com.itextpdf.layout import Document, Canvas
    from com.itextpdf.layout.layout import LayoutContext, LayoutArea
    from com.itextpdf.kernel.geom import Rectangle, PageSize
    from com.itextpdf.layout.properties import TextAlignment
    from com.itextpdf.layout.element import Paragraph, Table, Cell
    from com.itextpdf.kernel.pdf import PdfDocument, PdfWriter
    from com.itextpdf.kernel.pdf.xobject import PdfFormXObject
    from com.itextpdf.kernel.pdf.canvas import PdfCanvas
    from org.apache.log4j import PropertyConfigurator

    # Setup log4j logging environment
    PropertyConfigurator.configure("log4j.properties")

    createDoc()
    if jpype.isJVMStarted():
        print("Shutting down JVM")
        jpype.shutdownJVM()
