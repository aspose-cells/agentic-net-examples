// Title: Preserve leading apostrophe while applying bold style with StyleFlag in Aspose.Cells for .NET (C#)
// Common Searches: how to keep a leading apostrophe in an Excel cell when applying a style using Aspose.Cells C# | Aspose.Cells StyleFlag QuotePrefix false example for preserving cell text prefix | apply bold formatting to a cell without removing the quote character in .NET | C# code to style a cell while retaining the apostrophe prefix in Aspose.Cells | using StyleFlag to prevent stripping of leading single quote in Excel via Aspose.Cells
// Tags: StyleFlag QuotePrefix false Aspose.Cells | retain leading apostrophe during cell formatting | bold font style with preserved quote prefix | Aspose.Cells cell styling keep text prefix | C# Excel cell style without stripping apostrophe

using System;
using Aspose.Cells;

// The example creates a workbook, writes a value with a leading apostrophe into cell A1, defines a bold font style, sets a StyleFlag with FontBold true and QuotePrefix false to keep the apostrophe, applies the style to the cell range, and saves the file as StyledOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Example data that contains an apostrophe prefix
            worksheet.Cells["A1"].PutValue("'Sample Text");

            // Create a style (e.g., make the font bold)
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Configure a StyleFlag
            StyleFlag flag = new StyleFlag
            {
                FontBold = true,        // Apply the bold attribute
                QuotePrefix = false     // Preserve existing apostrophe (quote) prefixes
            };

            // Apply the style with the flag to the target cell
            Aspose.Cells.Range targetRange = worksheet.Cells.CreateRange("A1");
            targetRange.ApplyStyle(boldStyle, flag);

            // Save the workbook
            workbook.Save("StyledOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
