// Title: Preserve 3‑Color Scale Conditional Formatting When Exporting Excel to PDF with Aspose.Cells (C#)
// Description: The sample builds a workbook, fills column A with numeric values, adds a red‑yellow‑green three‑color‑scale conditional format to A1:A11, and saves the file as PDF using PdfSaveOptions (MergeAreas = true, CheckFontCompatibility = true, DefaultFont = "Arial"). The PDF output retains the conditional‑format colors exactly as shown in Excel.
// Keywords: Aspose.Cells PDF export | conditional formatting PDF | 3‑color scale Aspose | MergeAreas PDF | CheckFontCompatibility | C# Aspose.Cells example | Excel to PDF color scale | preserve conditional formatting | PdfSaveOptions | Aspose.Cells rendering options
// Common Searches: Aspose.Cells keep conditional formatting colors in PDF | PdfSaveOptions MergeAreas usage | Export Excel color scale to PDF C# | How to render conditional formatting in PDF with Aspose | Enable font compatibility Aspose PDF export
// Developer Intent: Export an Excel workbook to PDF while preserving the visual appearance of a three‑color‑scale conditional formatting.
// Use Cases: Generate printable financial reports where low, medium, and high values are highlighted with red, yellow, and green colors. | Create automated dashboards that retain Excel‑style color scaling when shared as PDF documents. | Produce compliance‑oriented PDFs where conditional formatting thresholds must match the original spreadsheet view.
// AI Prompts: Show C# code that saves an Aspose.Cells workbook to PDF with a 3‑color scale conditional format preserved, including the required PdfSaveOptions settings. | Explain why the MergeAreas property is necessary for conditional‑format colors to appear in PDF output using Aspose.Cells. | Give step‑by‑step instructions to enable font compatibility and set a default font when exporting Excel to PDF with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

// The sample builds a workbook, fills column A with numeric values, adds a red‑yellow‑green three‑color‑scale conditional format to A1:A11, and saves the file as PDF using PdfSaveOptions (MergeAreas = true, CheckFontCompatibility = true, DefaultFont = "Arial"). The PDF output retains the conditional‑format colors exactly as shown in Excel.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample numeric data
        for (int row = 0; row <= 10; row++)
        {
            worksheet.Cells[row, 0].PutValue(row * 10);
        }

        // Add a 3‑color scale conditional formatting to the data range
        int cfIndex = worksheet.ConditionalFormattings.Add();
        var cfCollection = worksheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A11
        var area = new CellArea { StartRow = 0, EndRow = 10, StartColumn = 0, EndColumn = 0 };
        cfCollection.AddArea(area);

        // Create the color‑scale condition
        int conditionIndex = cfCollection.AddCondition(FormatConditionType.ColorScale);
        var condition = cfCollection[conditionIndex];
        condition.ColorScale.Is3ColorScale = true;
        condition.ColorScale.MinColor = Color.Red;      // low values -> red
        condition.ColorScale.MidColor = Color.Yellow;   // middle values -> yellow
        condition.ColorScale.MaxColor = Color.Green;    // high values -> green

        // Configure PDF save options to preserve conditional formatting colors
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Merge conditional formatting areas before rendering (required for colors to appear)
            MergeAreas = true,

            // Ensure font compatibility so text is rendered correctly
            CheckFontCompatibility = true,
            DefaultFont = "Arial"
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("ConditionalFormattingColors.pdf", pdfOptions);
    }
}
