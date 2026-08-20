// Title: Export Each Worksheet to a Separate PDF with Aspose.Cells (C#) using PdfSaveOptions.SheetSet
// Description: The example builds a workbook with three sheets, writes a value to A1 on each sheet, then iterates through the worksheets. For every iteration it creates a PdfSaveOptions object, sets its SheetSet to the current sheet index, composes a file name that includes the sheet number and name, and saves that sheet as an individual PDF file.
// Keywords: Aspose.Cells | C# PDF export | PdfSaveOptions SheetSet | save worksheet as PDF | individual PDF per sheet | export Excel to PDF .NET | Aspose.Cells per‑sheet PDF | loop worksheets Aspose | Aspose.Cells PDF options | C# Aspose.Cells example
// Common Searches: Aspose.Cells export each worksheet to separate PDF | PdfSaveOptions SheetSet usage C# | How to save individual Excel sheets as PDF with Aspose | Loop through workbook worksheets and generate PDFs | C# Aspose.Cells save specific sheet to PDF
// Developer Intent: Generate a distinct PDF file for every worksheet in a workbook.
// Use Cases: Produce separate PDF reports for departmental tabs in a financial workbook. | Automate per‑sheet PDF delivery when each worksheet represents a client contract. | Archive each data‑collection sheet as an individual PDF for regulatory compliance.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over all worksheets and save each one as a PDF with a custom filename. | Show how to configure PdfSaveOptions.SheetSet to export only the current worksheet inside a loop. | Explain how to modify the sample to combine several non‑contiguous worksheets into a single PDF using SheetSet.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example builds a workbook with three sheets, writes a value to A1 on each sheet, then iterates through the worksheets. For every iteration it creates a PdfSaveOptions object, sets its SheetSet to the current sheet index, composes a file name that includes the sheet number and name, and saves that sheet as an individual PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook and add a few worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Fill each worksheet with sample data
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            ws.Cells["A1"].PutValue($"Data from {ws.Name}");
        }

        // Loop through each worksheet, set SheetSet to that sheet, and save as a separate PDF
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Configure SheetSet to include only the current worksheet (zero‑based index)
            pdfOptions.SheetSet = new SheetSet(new int[] { i });

            // Define output file name
            string outputFile = $"Sheet_{i + 1}_{workbook.Worksheets[i].Name}.pdf";

            // Save the selected sheet to PDF
            workbook.Save(outputFile, pdfOptions);
        }
    }
}
