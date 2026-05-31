using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Header");
            cells["B2"].PutValue(123);
            cells["C3"].PutValue("");          // empty string considered empty
            cells["D4"].PutValue("Data");

            // Retrieve the maximum display range (includes data, merged cells, shapes)
            Aspose.Cells.Range maxDisplayRange = cells.MaxDisplayRange;
            if (maxDisplayRange == null)
            {
                Console.WriteLine("The worksheet is empty.");
                return;
            }

            // Iterate through all cells in the range and count non‑empty cells
            int nonEmptyCount = 0;
            IEnumerator enumerator = maxDisplayRange.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Consider a cell non‑empty if its Value is not null and not an empty/whitespace string
                if (cell.Value != null && !(cell.Value is string str && string.IsNullOrWhiteSpace(str)))
                {
                    nonEmptyCount++;
                }
            }

            Console.WriteLine($"Non‑empty cells in MaxDisplayRange: {nonEmptyCount}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "NonEmptyCountDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}