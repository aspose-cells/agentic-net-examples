// Title: Generate a dynamic G3:H range up to the last used row and apply bold font using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a range from G3 to the last populated row covering columns G and H, then applies a bold font style using a StyleFlag. | Show how to find the last used row in a worksheet, build a multi‑column range with CreateRange, and set only the FontBold attribute without altering other cell styles.
// Common Searches: Aspose.Cells C# create range from G3 to last row in columns G and H | How to apply only bold formatting to a range in Aspose.Cells .NET | C# Aspose.Cells dynamic column range styling with StyleFlag | Set bold font for G3:H range based on last used row using Aspose.Cells
// Tags: Aspose.Cells create dynamic range G3:H | Aspose.Cells apply bold style with StyleFlag | C# determine last used row Aspose.Cells | Aspose.Cells multi‑column range formatting | Aspose.Cells save workbook after styling

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, determines the last used row, builds a range covering columns G and H from row 3 to that row, applies a bold font style using a StyleFlag, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the start of the range (row 3, column G)
            int startRow = 2;          // zero‑based index for row 3
            int startColumn = 6;       // zero‑based index for column G

            // Determine the number of rows from the start row to the last used row
            int totalRows = sheet.Cells.MaxRow - startRow + 1;
            int totalColumns = 2; // Columns G and H

            // Create the range representing G3:H{last row}
            AsposeRange range = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Prepare a style with bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Apply the bold style only to the font attribute
            StyleFlag flag = new StyleFlag { FontBold = true };
            range.ApplyStyle(boldStyle, flag);

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
