// Title: Add a hyperlink in Aspose.Cells for .NET that opens a specific cell in another Excel workbook
// Common Searches: Aspose.Cells C# add hyperlink to cell B5 in external workbook Book2.xlsx | How to use external reference syntax for hyperlinks with Aspose.Cells .NET | Create a blue underlined hyperlink in Excel using Aspose.Cells C# | Set hyperlink style programmatically with Aspose.Cells | Save workbook containing external cell link using Aspose.Cells
// Tags: hyperlink to external workbook cell Aspose.Cells | Aspose.Cells external hyperlink reference | hyperlink visual style blue Aspose.Cells | programmatic hyperlink formatting Aspose.Cells | saving workbook with external cell link Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// This C# example creates a new workbook, inserts a hyperlink in cell A1 that points to cell B5 on Sheet1 of an external file 'Book2.xlsx' using Aspose.Cells external reference syntax, applies a blue underlined style to mimic a typical hyperlink, and saves the file as 'HyperlinkExternalReference.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell where the hyperlink will be placed (A1)
            int firstRow = 0;      // Row index (0‑based)
            int firstColumn = 0;   // Column index (0‑based)

            // Text that will be displayed in the cell
            string displayText = "Open External Cell";

            // External reference syntax: '[WorkbookName]SheetName'!CellAddress
            // Example: link to cell B5 in Sheet1 of Book2.xlsx
            string externalReference = "'[Book2.xlsx]Sheet1'!B5";

            // Add the hyperlink to the worksheet.
            // Correct overload order: address, firstRow, firstColumn, displayText
            sheet.Hyperlinks.Add(externalReference, firstRow, firstColumn, displayText);

            // Optionally, set the cell style to look like a hyperlink (blue & underlined)
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.Blue;
            style.Font.Underline = FontUnderlineType.Single;

            // Apply the style to the target cell
            sheet.Cells[firstRow, firstColumn].SetStyle(style);

            // Save the workbook
            string outputPath = "HyperlinkExternalReference.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
