// Title: Aspose.Cells .NET: Export Worksheet to PDF Before and After AutoFitRows and Compare File Sizes
// Description: C# example that creates a workbook, adds wrapped long text, saves a PDF with default row heights, applies Worksheet.AutoFitRows(), saves a second PDF, and outputs the byte sizes to show how AutoFitRows impacts PDF file size.
// Keywords: Aspose.Cells AutoFitRows PDF | C# export Excel to PDF | compare PDF size Aspose.Cells | row auto fit PDF output .NET | worksheet.AutoFitRows effect on file size
// Common Searches: Aspose.Cells export PDF before AutoFitRows | Does AutoFitRows increase PDF size in .NET | How to compare PDF file sizes with and without row auto‑fit | C# Aspose.Cells PDF size optimization | Measure impact of AutoFitRows on PDF output
// Developer Intent: Generate two PDFs—one with default row heights and one after applying AutoFitRows—and determine the size difference.
// Use Cases: Validate that row auto‑fitting does not unnecessarily enlarge PDF reports. | Choose the optimal PDF generation strategy based on file‑size impact. | Integrate a CI check that flags PDFs whose size grows after AutoFitRows.
// AI Prompts: Write C# code using Aspose.Cells to save a worksheet as PDF, apply AutoFitRows, save again, and report the size delta. | Create a method that accepts a Workbook, exports PDFs before and after AutoFitRows, and returns true if the second PDF is larger. | Explain how text wrapping and AutoFitRows influence PDF rendering and file size in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a workbook, adds wrapped long text, saves a PDF with default row heights, applies Worksheet.AutoFitRows(), saves a second PDF, and outputs the byte sizes to show how AutoFitRows impacts PDF file size.
class AutoFitRowsPdfComparison
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with long text to demonstrate row height changes
        worksheet.Cells["A1"].PutValue("This is a very long text that will require row auto‑fitting. It contains multiple sentences to increase the height.");
        worksheet.Cells["A2"].PutValue("Short text");
        worksheet.Cells["A3"].PutValue("Another long text that should wrap and increase row height when auto‑fit is applied.");

        // Enable text wrapping for the cells that contain long text
        Style wrapStyle = worksheet.Cells["A1"].GetStyle();
        wrapStyle.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(wrapStyle);
        worksheet.Cells["A3"].SetStyle(wrapStyle);

        // Save PDF without applying AutoFitRows
        string pdfWithoutFit = "WithoutAutoFit.pdf";
        workbook.Save(pdfWithoutFit, SaveFormat.Pdf);

        // AutoFit all rows in the worksheet
        worksheet.AutoFitRows();

        // Save PDF after applying AutoFitRows
        string pdfWithFit = "WithAutoFit.pdf";
        workbook.Save(pdfWithFit, SaveFormat.Pdf);

        // Compare file sizes
        long sizeWithoutFit = new FileInfo(pdfWithoutFit).Length;
        long sizeWithFit = new FileInfo(pdfWithFit).Length;

        Console.WriteLine($"PDF size without AutoFitRows: {sizeWithoutFit} bytes");
        Console.WriteLine($"PDF size with AutoFitRows: {sizeWithFit} bytes");
        Console.WriteLine(sizeWithFit > sizeWithoutFit
            ? "AutoFitRows increased the PDF size."
            : "AutoFitRows reduced or kept the PDF size unchanged.");
    }
}
