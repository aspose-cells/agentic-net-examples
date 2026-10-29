// Title: Create a blue underlined hyperlink to an external PDF that opens in a new browser tab using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add hyperlink to external PDF that opens in new tab | How to set display text for a PDF link in an Excel cell with Aspose.Cells | Formatting Excel cell as hyperlink with blue underline using Aspose.Cells .NET | Save Excel workbook with PDF hyperlink using Aspose.Cells API | Add single‑cell hyperlink to PDF URL in Aspose.Cells C# example
// Tags: Aspose.Cells external PDF link creation | C# Excel hyperlink cell formatting | display text for Aspose.Cells hyperlink | persist workbook with PDF hyperlink | blue underlined hyperlink style Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new workbook, accesses the first worksheet, adds a hyperlink in cell A1 that points to an external PDF URL, sets the cell value to "Open PDF" as the display text, applies a blue underlined style to mimic a typical hyperlink, and saves the workbook as HyperlinkToPdf.xlsx.
class HyperlinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell where the hyperlink will be placed (e.g., A1)
            int row = 0;    // zero‑based index for row 1
            int column = 0; // zero‑based index for column A

            // Add a hyperlink to the specified cell (single cell range)
            // Note: Aspose.Cells does not have an overload that sets display text directly.
            // Set the cell value to the desired display text after adding the hyperlink.
            sheet.Hyperlinks.Add(row, column, 1, 1, "https://example.com/document.pdf");
            sheet.Cells[row, column].PutValue("Open PDF");

            // Optionally, format the cell to look like a typical hyperlink
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.Blue;
            style.Font.Underline = FontUnderlineType.Single;
            sheet.Cells[row, column].SetStyle(style);

            // Save the workbook to a file
            string outputPath = "HyperlinkToPdf.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
