// +build windows

// main_windows
package main

import (
    "doc2pdf/office2pdf"
    "fmt"
    "log"
    "os"
    "path/filepath"
)

func fileIsExist(path string) bool {
    if _, err := os.Stat(path); os.IsNotExist(err) {
        return false
    }
    return true
}

func exporterMap() (m map[string]interface{}) {
    m = map[string]interface{}{
        ".doc":  new(office2pdf.Word),
        ".docx": new(office2pdf.Word),
        ".xls":  new(office2pdf.Excel),
        ".xlsx": new(office2pdf.Excel),
        ".ppt":  new(office2pdf.PowerPoint),
        ".pptx": new(office2pdf.PowerPoint),
    }
    return
}

func main() {
    inFile, outDir := os.Args[1], os.Args[2]
    if len(os.Args) > 2 && fileIsExist(inFile) && fileIsExist(outDir) {
        exporter := exporterMap()[filepath.Ext(inFile)]
        if _, ok := exporter.(office2pdf.Exporter); ok {
            outFile, err := exporter.(office2pdf.Exporter).Export(inFile, outDir)
            if err != nil {
                log.Fatal(err)
            }
            fmt.Printf("%v", outFile)
        }
    }
}
