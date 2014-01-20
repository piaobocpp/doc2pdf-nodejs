package main

import (
    "fmt"
    "github.com/mattn/go-ole"
    "github.com/mattn/go-ole/oleutil"
    "log"
    "os"
    "path/filepath"
)

func wordToPDF(inFile, outDir string) (outFile string, err error) {

    ole.CoInitialize(0)

    var (
        unknown   *ole.IUnknown
        app       *ole.IDispatch
        documents *ole.VARIANT
        doc       *ole.VARIANT
    )

    outFile = filepath.Join(outDir, filepath.Base(inFile+".pdf"))

    defer func() {
        if err != nil {
            outFile = ""
        }

        if doc != nil {
            oleutil.PutProperty(doc.ToIDispatch(), "Saved", true)
        }

        if documents != nil {
            oleutil.CallMethod(documents.ToIDispatch(), "Close")
        }

        if app != nil {
            oleutil.MustCallMethod(app, "Quit")
            app.Release()
        }

        ole.CoUninitialize()
    }()

    unknown, err = oleutil.CreateObject("Word.Application")
    if err != nil {
        return
    }

    app, err = unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(app, "Visible", false)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(app, "DisplayAlerts", 0)
    if err != nil {
        return
    }

    documents, err = oleutil.GetProperty(app, "Documents")
    if err != nil {
        return
    }

    doc, err = oleutil.CallMethod(documents.ToIDispatch(), "Open", inFile)
    if err != nil {
        return
    }

    _, err = oleutil.CallMethod(doc.ToIDispatch(), "ExportAsFixedFormat", outFile, 17)
    if err != nil {
        return
    }

    return
}

func excelToPDF(inFile, outDir string) (outFile string, err error) {

    ole.CoInitialize(0)

    var (
        unknown   *ole.IUnknown
        app       *ole.IDispatch
        workbooks *ole.VARIANT
        xls       *ole.VARIANT
    )

    outFile = filepath.Join(outDir, filepath.Base(inFile+".pdf"))

    defer func() {
        if err != nil {
            outFile = ""
        }

        if xls != nil {
            oleutil.PutProperty(xls.ToIDispatch(), "Saved", true)
        }

        if workbooks != nil {
            oleutil.CallMethod(workbooks.ToIDispatch(), "Close")
        }

        if app != nil {
            oleutil.MustCallMethod(app, "Quit")
            app.Release()
        }

        ole.CoUninitialize()
    }()

    unknown, err = oleutil.CreateObject("Excel.Application")
    if err != nil {
        return
    }

    app, err = unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(app, "Visible", false)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(app, "DisplayAlerts", false)
    if err != nil {
        return
    }

    workbooks, err = oleutil.GetProperty(app, "Workbooks")
    if err != nil {
        return
    }

    xls, err = oleutil.CallMethod(workbooks.ToIDispatch(), "Open", inFile)
    if err != nil {
        return
    }

    _, err = oleutil.CallMethod(xls.ToIDispatch(), "ExportAsFixedFormat", 0, outFile)
    if err != nil {
        return
    }

    return
}

func powerpointToPDF(inFile, outDir string) (outFile string, err error) {

    ole.CoInitialize(0)

    var (
        unknown       *ole.IUnknown
        app           *ole.IDispatch
        presentations *ole.VARIANT
        ppt           *ole.VARIANT
    )

    outFile = filepath.Join(outDir, filepath.Base(inFile+".pdf"))

    defer func() {
        if err != nil {
            outFile = ""
        }

        if ppt != nil {
            oleutil.PutProperty(ppt.ToIDispatch(), "Saved", -1)
            oleutil.MustCallMethod(ppt.ToIDispatch(), "Close")
        }

        if app != nil {
            oleutil.MustCallMethod(app, "Quit")
            app.Release()
        }

        ole.CoUninitialize()
    }()

    unknown, err = oleutil.CreateObject("PowerPoint.Application")
    if err != nil {
        return
    }

    app, err = unknown.QueryInterface(ole.IID_IDispatch)
    if err != nil {
        return
    }

    _, err = oleutil.PutProperty(app, "DisplayAlerts", 1)
    if err != nil {
        return
    }

    presentations, err = oleutil.GetProperty(app, "Presentations")
    if err != nil {
        return
    }

    ppt, err = oleutil.CallMethod(presentations.ToIDispatch(), "Open", inFile, -1, 0, 0)
    if err != nil {
        return
    }

    _, err = oleutil.CallMethod(ppt.ToIDispatch(), "SaveAs", outFile, 32)
    if err != nil {
        return
    }

    return
}

func fileIsExist(path string) bool {
    if _, err := os.Stat(path); os.IsNotExist(err) {
        return false
    }
    return true
}

func main() {
    inFile, outDir := os.Args[1], os.Args[2]
    if len(os.Args) > 2 && fileIsExist(inFile) && fileIsExist(outDir) {
        switch filepath.Ext(inFile) {
        case ".doc":
            fallthrough
        case ".docx":
            outFile, err := wordToPDF(inFile, outDir)
            if err != nil {
                log.Fatal(err)
            }
            fmt.Printf("%v", outFile)
        case ".xls":
            fallthrough
        case ".xlsx":
            outFile, err := excelToPDF(inFile, outDir)
            if err != nil {
                log.Fatal(err)
            }
            fmt.Printf("%v", outFile)
        case ".ppt":
            fallthrough
        case ".pptx":
            outFile, err := powerpointToPDF(inFile, outDir)
            if err != nil {
                log.Fatal(err)
            }
            fmt.Printf("%v", outFile)
        }
    }
}
