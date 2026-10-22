// Title: Create a G3:H column range to the last used row and apply bold font with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create range from G3 to the last row in columns G and H and set bold font | How to apply bold style to a dynamic column range using Aspose.Cells in .NET | Aspose.Cells example for formatting entire columns starting at row 3 | Apply StyleFlag FontBold to a multi‑column range in an Excel workbook with Aspose.Cells | Generate a range covering G3:H and make text bold in a .xlsx file using Aspose.Cells
// Tags: create range G3:H Aspose.Cells | bold font StyleFlag Aspose.Cells | dynamic range to last used row Aspose.Cells | format columns G and H Aspose.Cells | C# Aspose.Cells range styling

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample creates a new workbook, defines a range that starts at cell G3 and extends through column H to the worksheet's last used row, builds a bold font style, applies it with a StyleFlag that targets only the font bold attribute, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last row index (zero‑based)
                int lastRowIndex = sheet.Cells.MaxRow;

                // Define range start at G3 (row index 2, column index 6)
                int startRow = 2;          // G3 row (zero‑based)
                int startColumn = 6;       // column G
                int totalRows = lastRowIndex - startRow + 1; // rows from G3 to last row
                int totalColumns = 2;      // columns G and H

                // Create the range covering G3:H(last row)
                Aspose.Cells.Range range = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

                // Create a style with bold font
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;

                // Apply only the bold font attribute
                StyleFlag flag = new StyleFlag { FontBold = true };
                range.ApplyStyle(boldStyle, flag);

                // Prepare output file path
                string fileName = "BoldColumns_G3_H_LastRow.xlsx";
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
