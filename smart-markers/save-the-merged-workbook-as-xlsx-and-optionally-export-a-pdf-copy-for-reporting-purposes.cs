// Title: Save a merged workbook as XLSX and optionally generate a PDF report with Aspose.Cells for .NET
// Common Searches: how to save a combined workbook to xlsx with Aspose.Cells in C# | conditional PDF generation from an Aspose.Cells workbook in .NET | export merged spreadsheet to pdf only when needed using Aspose.Cells | Aspose.Cells C# example for creating a PDF copy of an existing workbook | save workbook in both xlsx and pdf formats with Aspose.Cells
// Tags: Aspose.Cells generate xlsx from combined workbook | Aspose.Cells optional pdf output | Aspose.Cells SaveFormat Xlsx Pdf example | C# export workbook to pdf using Aspose.Cells | smart markers workbook export workflow

using Aspose.Cells;
using System;

// Demonstrates obtaining a merged Workbook object, saving it as an XLSX file, and, based on a boolean flag, optionally exporting the same workbook to PDF for reporting purposes using Aspose.Cells for .NET.
class WorkbookExporter
{
    static void Main()
    {
        // Assume the merged workbook is already created and populated
        Workbook workbook = GetMergedWorkbook();

        // Save the workbook as XLSX
        workbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);

        // Optionally export a PDF copy for reporting purposes
        bool exportPdf = true; // Set to false if PDF export is not needed
        if (exportPdf)
        {
            workbook.Save("MergedWorkbook.pdf", SaveFormat.Pdf);
        }
    }

    // Placeholder method to obtain the merged workbook.
    // Replace with actual logic that creates or loads the merged workbook.
    static Workbook GetMergedWorkbook()
    {
        // Example: create a new workbook (replace with real merged workbook)
        Workbook wb = new Workbook();
        // ... populate workbook ...
        return wb;
    }
}
