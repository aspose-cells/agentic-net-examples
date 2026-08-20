// Title: Load Excel workbook with Letter paper size, evaluate PDF page count, and save as PDF using Aspose.Cells for .NET
// Description: Demonstrates how to set the default paper size to Letter when loading an Excel file, retrieve the evaluated page count via WorkbookPrintingPreview, and export the workbook to PDF while ensuring the generated PDF matches the expected pagination.
// Keywords: Aspose.Cells .NET | set paper size Letter | LoadOptions paper size | WorkbookPrintingPreview page count | evaluate PDF pagination | Excel to PDF conversion | ImageOrPrintOptions | SaveFormat.Pdf
// Common Searches: Aspose.Cells set default paper size to Letter | How to get evaluated page count before PDF export | Match PDF page count with workbook preview Aspose.Cells | Load Excel with custom paper size .NET | WorkbookPrintingPreview usage example
// Developer Intent: Determine the number of pages an Excel workbook will occupy on Letter paper and generate a PDF that reflects that exact pagination.
// Use Cases: Create printable PDF reports that conform to US Letter dimensions. | Validate pagination of workbooks prior to batch PDF conversion. | Automate document workflows where page count consistency is required for legal or archival purposes.
// AI Prompts: Generate C# code that loads an Excel file with Letter paper size and returns the evaluated PDF page count using Aspose.Cells. | Explain the role of WorkbookPrintingPreview in calculating pagination and how to compare its result with the final PDF output. | Suggest strategies for handling discrepancies between evaluated and actual PDF page counts during conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to set the default paper size to Letter when loading an Excel file, retrieve the evaluated page count via WorkbookPrintingPreview, and export the workbook to PDF while ensuring the generated PDF matches the expected pagination.
class Program
{
    static void Main()
    {
        // Set load options to use Letter paper size as default
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.SetPaperSize(PaperSizeType.PaperLetter);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Ensure the workbook's default paper size is Letter
        workbook.Settings.PaperSize = PaperSizeType.PaperLetter;

        // Create print options (default settings are sufficient for PDF rendering)
        ImageOrPrintOptions printOptions = new ImageOrPrintOptions();

        // Evaluate total page count for the workbook
        WorkbookPrintingPreview preview = new WorkbookPrintingPreview(workbook, printOptions);
        int evaluatedPageCount = preview.EvaluatedPageCount;
        Console.WriteLine($"Evaluated page count: {evaluatedPageCount}");

        // Save the workbook as PDF; the page count should match the evaluated count
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
