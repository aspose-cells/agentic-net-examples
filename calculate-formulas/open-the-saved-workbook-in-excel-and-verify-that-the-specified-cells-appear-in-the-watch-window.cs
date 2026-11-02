// Title: C# Aspose.Cells example: verify that specific cells contain data in an existing .xlsx workbook
// Common Searches: C# Aspose.Cells check if cell A1 has a value in an existing Excel file | How to validate non‑empty cells B2 and C3 using Aspose.Cells for .NET | Programmatically confirm presence of specific cells before opening Excel with Aspose.Cells | Aspose.Cells .NET read cell addresses and test for empty values | Load workbook with Aspose.Cells and verify cell data in C#
// Tags: Aspose.Cells load workbook .xlsx | Aspose.Cells read cell value C# | Aspose.Cells check cell emptiness | Aspose.Cells verify specific cell addresses | Aspose.Cells handle missing file exception

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing .xlsx workbook using Aspose.Cells for .NET, accesses the first worksheet, iterates over the cell addresses A1, B2, and C3, and reports whether each cell contains a non‑empty value, with proper file‑existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file
            string filePath = @"C:\Temp\MyWorkbook.xlsx";

            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(filePath);

                // Get the first worksheet (or any specific worksheet you need)
                Worksheet sheet = workbook.Worksheets[0];

                // Define the cell addresses you expect to find
                string[] expectedCells = new string[] { "A1", "B2", "C3" };

                // Verify each expected cell is present (has a non‑empty value)
                foreach (string address in expectedCells)
                {
                    bool found = false;

                    // Retrieve the cell by its address
                    Cell cell = sheet.Cells[address];

                    // Consider the cell present if it contains any value
                    if (cell != null && cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        found = true;
                    }

                    // Output verification result
                    if (found)
                    {
                        Console.WriteLine($"Cell {address} is present and contains data.");
                    }
                    else
                    {
                        Console.WriteLine($"Cell {address} is NOT present or empty.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
