// Title: Convert an Excel ListObject to a regular range while preserving formatting for the first ten rows using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert ListObject to range keep first 10 rows formatting | how to retain cell styles for the first ten rows after ConvertToRange in Aspose.Cells | C# Aspose.Cells example clearing formatting beyond a specific row in a converted table | convert Excel table to range while preserving header formatting with Aspose.Cells .NET | Aspose.Cells remove styles after row 10 in a ListObject conversion
// Tags: ListObject to range conversion Aspose.Cells | retain first ten rows formatting Aspose.Cells | remove styles after row ten Aspose.Cells | Excel table conversion C# Aspose.Cells | partial formatting preservation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an Excel workbook (or creates one), accesses a ListObject named "Table1", converts it to a normal range with ConvertToRange(), clears cell styles for rows beyond the tenth row while keeping the original formatting for the first ten rows, and saves the result.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook or create a sample one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                // Sample data for the table
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                for (int i = 0; i < 20; i++)
                {
                    ws.Cells[i + 1, 0].PutValue($"R{i + 1}C1");
                    ws.Cells[i + 1, 1].PutValue($"R{i + 1}C2");
                }

                // Add a ListObject (table) named "Table1"
                int tableIndex = ws.ListObjects.Add(0, 0, 20, 1, true);
                ListObject table = ws.ListObjects[tableIndex];
                table.DisplayName = "Table1";

                // Save the generated input file for reference
                workbook.Save(inputPath);
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the table by name
            ListObject tableObj = worksheet.ListObjects["Table1"];
            if (tableObj == null)
            {
                Console.WriteLine("Table 'Table1' not found.");
                return;
            }

            // Preserve formatting information before converting to a normal range
            int startRow = tableObj.StartRow;
            int startColumn = tableObj.StartColumn;
            int endRow = tableObj.EndRow;
            int endColumn = tableObj.EndColumn;

            // Convert the table to a normal range while keeping values and formatting
            tableObj.ConvertToRange();

            // Keep formatting only for the first 10 rows of the original table
            int formattingEndRow = startRow + 9; // 10th row (0‑based)

            // Create a default (no) style
            Style defaultStyle = workbook.CreateStyle();

            // Clear formatting for rows beyond the 10th row within the former table area
            for (int r = formattingEndRow + 1; r <= endRow; r++)
            {
                for (int c = startColumn; c <= endColumn; c++)
                {
                    worksheet.Cells[r, c].SetStyle(defaultStyle);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
