// Title: Convert a CSV file to PDF with column auto‑fit and single‑page width using Aspose.Cells in C#
// AI Prompts: Write C# code that loads a CSV file with Aspose.Cells, automatically adjusts column widths, configures the worksheet to fit the entire width on one PDF page, and saves the workbook as a PDF. | Explain how to set Aspose.Cells page‑setup properties so that exporting a CSV‑derived worksheet to PDF keeps all columns within a single page width.
// Common Searches: Aspose.Cells C# auto adjust column widths when converting CSV to PDF | fit worksheet columns to one page width in PDF export using Aspose.Cells | load CSV with LoadOptions and preserve column alignment in PDF with Aspose.Cells .NET | example of setting page setup for single‑page PDF output in Aspose.Cells
// Tags: CSV to PDF conversion Aspose.Cells C# | adjust column widths Aspose.Cells worksheet | single‑page width PDF export Aspose.Cells | LoadOptions CSV Aspose.Cells .NET | preserve column layout PDF conversion Aspose.Cells

using Aspose.Cells;
using System;

// Loads a CSV file into an Aspose.Cells workbook, automatically adjusts column widths, configures the page setup to fit all columns on one PDF page, and saves the result as a PDF.
class Program
{
    static void Main()
    {
        // Input CSV file path
        string csvPath = "input.csv";

        // Output PDF file path
        string pdfPath = "output.pdf";

        // Load the CSV file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Work with the first worksheet (the CSV data is loaded here)
        Worksheet sheet = workbook.Worksheets[0];

        // Auto‑fit all columns so the width matches the source data
        sheet.AutoFitColumns();

        // Configure page setup to keep all columns on a single page width
        PageSetup pageSetup = sheet.PageSetup;
        pageSetup.FitToPagesWide = 1;   // fit all columns to one page width
        pageSetup.FitToPagesTall = 0;   // allow any number of pages tall

        // Save the workbook as a PDF file
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }
}
