"""Gera scripts/vba/vbaProject.bin a partir de scripts/vba/ModAcerto.bas usando o LibreOffice.

Uso:
    soffice --headless --accept="socket,host=localhost,port=2002;urp;" --norestore &
    python3 scripts/gerar_vba.py <numero_de_abas>

O vbaProject.bin resultante contem os modulos de documento ThisWorkbook e
Sheet1..SheetN (os codenames usados pelo gerar_planilha.py) e o modulo ModAcerto.
"""
import os
import sys

import uno
from com.sun.star.beans import PropertyValue

AQUI = os.path.dirname(os.path.abspath(__file__))


def prop(nome, valor):
    p = PropertyValue()
    p.Name = nome
    p.Value = valor
    return p


def main():
    n_abas = int(sys.argv[1])
    codigo = open(os.path.join(AQUI, "vba", "ModAcerto.bas"), encoding="ascii").read()

    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, (prop("Hidden", True),))

    sheets = doc.Sheets
    while sheets.Count < n_abas:
        sheets.insertNewByName("S%d" % (sheets.Count + 1), sheets.Count)
    for i in range(n_abas):
        sheets.getByIndex(i).CodeName = "Sheet%d" % (i + 1)

    libs = doc.BasicLibraries
    libs.VBACompatibilityMode = True
    libs.ProjectName = "VBAProject"
    lib = libs.createLibrary("VBAProject")
    ModuleInfo = uno.getClass("com.sun.star.script.ModuleInfo")

    def add(nome, src, tipo):
        mi = ModuleInfo()
        mi.ModuleType = tipo
        lib.insertModuleInfo(nome, mi)
        lib.insertByName(nome, src)

    doc_hdr = "Rem Attribute VBA_ModuleType=VBADocumentModule\nOption VBASupport 1\n"
    add("ThisWorkbook", doc_hdr, 4)
    for i in range(n_abas):
        add("Sheet%d" % (i + 1), doc_hdr, 4)
    add("ModAcerto", "Rem Attribute VBA_ModuleType=VBAModule\nOption VBASupport 1\n" + codigo, 1)

    tmp = os.path.join(AQUI, "vba", "_tmp.xlsm")
    doc.storeToURL(uno.systemPathToFileUrl(tmp), (prop("FilterName", "Calc MS Excel 2007 VBA XML"),))
    doc.close(True)

    import zipfile
    with zipfile.ZipFile(tmp) as z:
        data = z.read("xl/vbaProject.bin")
    with open(os.path.join(AQUI, "vba", "vbaProject.bin"), "wb") as f:
        f.write(data)
    os.remove(tmp)
    print("vbaProject.bin gerado (%d bytes)" % len(data))


if __name__ == "__main__":
    main()
