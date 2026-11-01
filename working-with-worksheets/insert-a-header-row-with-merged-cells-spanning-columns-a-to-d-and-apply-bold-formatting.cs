// Title: Create a merged header row across columns A‑D with bold centered text using Aspose.Cells for .NET (C#)
// Common Searches: how to merge cells A1 to D1 and set bold centered style with Aspose.Cells in C# | Aspose.Cells C# example for creating a header row with merged cells and bold formatting | C# code to apply bold and center alignment to a merged range in an Excel worksheet using Aspose.Cells | using Aspose.Cells to merge a row and style it as a header in .NET | Aspose.Cells .NET merge cells and apply style flag for bold centered header
// Tags: Aspose.Cells merge range across columns | Aspose.Cells apply bold font style | Aspose.Cells center alignment for merged cells | Excel header row styling with Aspose.Cells .NET | Aspose.Cells style flag for cell formatting

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

// The program creates a new workbook, merges cells A1 through D1 on the first worksheet, inserts the text "Header Title", applies a bold centered style to the merged range using a StyleFlag, and saves the workbook as Output.xlsx.
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

            // Insert header text in cell A1
            Cell headerCell = sheet.Cells["A1"];
            headerCell.PutValue("Header Title");

            // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
            sheet.Cells.Merge(0, 0, 1, 4);

            // Create a style with bold font and centered alignment
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            style.HorizontalAlignment = TextAlignmentType.Center;

            // Define which style attributes to apply
            StyleFlag flag = new StyleFlag
            {
                FontBold = true,
                HorizontalAlignment = true
            };

            // Apply the style to the merged range A1:D1
            AsposeRange headerRange = sheet.Cells.CreateRange("A1:D1");
            headerRange.ApplyStyle(style, flag);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
