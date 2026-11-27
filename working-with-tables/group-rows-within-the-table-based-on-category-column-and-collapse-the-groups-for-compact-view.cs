// Title: How to group and collapse rows by a Category column in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# code to group consecutive rows with the same value and hide the group | programmatically create outline groups in Excel based on a column using Aspose.Cells | C# example for grouping rows by category and collapsing them in an .xlsx file | using Aspose.Cells to auto‑group duplicate rows in a spreadsheet | how to apply row grouping and collapse in a generated Excel workbook with Aspose.Cells
// Tags: Aspose.Cells outline rows programmatically | collapse grouped rows Excel .NET | auto‑group rows based on column value | C# row grouping with Cells.GroupRows | generate Excel outline using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample loads (or creates) an Excel workbook, scans the Category column, groups consecutive rows that share the same category using Cells.GroupRows, collapses each group for a compact outline view, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; create a simple one if it does not.
                if (!File.Exists(inputPath))
                {
                    var wb = new Workbook();
                    var ws = wb.Worksheets[0];
                    ws.Cells[0, 0].PutValue("Category");
                    ws.Cells[1, 0].PutValue("A");
                    ws.Cells[2, 0].PutValue("A");
                    ws.Cells[3, 0].PutValue("B");
                    ws.Cells[4, 0].PutValue("B");
                    ws.Cells[5, 0].PutValue("C");
                    wb.Save(inputPath);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                int categoryColumnIndex = 0;          // Column A (0‑based)
                int firstDataRow = 1;                 // Data starts after header row.
                int lastDataRow = sheet.Cells.MaxDataRow;

                int groupStart = firstDataRow;
                string previousCategory = null;

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    string currentCategory = sheet.Cells[row, categoryColumnIndex].StringValue;

                    if (previousCategory == null)
                    {
                        previousCategory = currentCategory;
                        continue;
                    }

                    // Category changed – close the previous group if it has more than one row.
                    if (!string.Equals(currentCategory, previousCategory, StringComparison.OrdinalIgnoreCase))
                    {
                        if (row - 1 > groupStart)
                        {
                            int totalRows = (row - 1) - groupStart + 1;
                            sheet.Cells.GroupRows(groupStart, totalRows, true);
                        }

                        groupStart = row;
                        previousCategory = currentCategory;
                    }
                }

                // Handle the final group.
                if (lastDataRow > groupStart)
                {
                    int totalRows = lastDataRow - groupStart + 1;
                    sheet.Cells.GroupRows(groupStart, totalRows, true);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
