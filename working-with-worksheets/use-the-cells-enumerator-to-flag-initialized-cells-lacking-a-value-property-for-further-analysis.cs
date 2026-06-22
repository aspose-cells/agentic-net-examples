using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEnumeratorExample
{
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

                // Populate some cells – some with values, some left empty (initialized but no value)
                cells["A1"].PutValue("Hello");
                cells["B1"].PutValue(123);
                _ = cells["C1"]; // Initialized cell without a value (discarded)
                _ = cells["A2"]; // Initialized cell without a value (discarded)
                cells["B2"].PutValue(DateTime.Now);
                cells["C2"].PutValue(null); // Explicitly set to null (treated as no value)

                // List to hold addresses of cells that are initialized but lack a value
                List<string> flaggedCells = new List<string>();

                // Use the Cells.GetEnumerator method to iterate through all instantiated cells
                IEnumerator enumerator = cells.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;
                    // Cell is instantiated; check if its Value is null (no data)
                    if (cell != null && cell.Value == null)
                    {
                        flaggedCells.Add(cell.Name);
                    }
                }

                // Output the flagged cells for further analysis
                Console.WriteLine("Initialized cells without a value:");
                foreach (string address in flaggedCells)
                {
                    Console.WriteLine(address);
                }

                // Save the workbook (optional, demonstrates lifecycle compliance)
                string outputPath = "EnumeratorFlaggedCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}