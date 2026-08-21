// Title: Set a DisplayRange with CellArea in Aspose.Cells for .NET and enumerate its cells
// Description: Creates a workbook, fills A1:C3 with sample data, defines a visible area using CellArea, builds a Range object from that area, iterates only the cells inside the range, outputs each cell's address and value, and saves the file.
// Keywords: Aspose.Cells .NET display range | CellArea CreateCellArea example | enumerate cells in Aspose.Cells range | visible cells iteration C# | Aspose.Cells Range object usage | filter cells by display area
// Common Searches: how to set display range Aspose.Cells .NET | enumerate cells inside a specific range Aspose.Cells | CellArea to define visible area in Aspose.Cells | iterate over a subset of worksheet cells C# | save workbook after defining display range Aspose.Cells
// Developer Intent: The developer wants to define a specific display area in a worksheet and loop through only the cells that belong to that area.
// Use Cases: Process a printable block of data without scanning the entire sheet. | Extract values from a known data region in a large workbook efficiently. | Generate a custom report by iterating over a subset of cells defined by CellArea.
// AI Prompts: Give C# code that sets a DisplayRange from B2 to D5 in Aspose.Cells and lists each cell's address and value. | Show how to create a CellArea, convert it to a Range, and enumerate only visible cells while ignoring hidden rows or columns. | Explain how to combine a defined DisplayRange with conditional formatting using Aspose.Cells.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDisplayRangeDemo
{
    // Creates a workbook, fills A1:C3 with sample data, defines a visible area using CellArea, builds a Range object from that area, iterates only the cells inside the range, outputs each cell's address and value, and saves the file.
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

                // Populate sample data (A1:C3)
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["C1"].PutValue("Header3");
                cells["A2"].PutValue(10);
                cells["B2"].PutValue(20);
                cells["C2"].PutValue(30);
                cells["A3"].PutValue(40);
                cells["B3"].PutValue(50);
                cells["C3"].PutValue(60);

                // Define the display range (rows 0‑2, columns 0‑2) using CellArea
                CellArea displayArea = CellArea.CreateCellArea(0, 0, 2, 2);

                // Create a Range object that corresponds to the display area
                int totalRows = displayArea.EndRow - displayArea.StartRow + 1;
                int totalCols = displayArea.EndColumn - displayArea.StartColumn + 1;

                // Resolve ambiguity with System.Range by using fully qualified name
                Aspose.Cells.Range visibleRange = cells.CreateRange(
                    displayArea.StartRow,
                    displayArea.StartColumn,
                    totalRows,
                    totalCols);

                // Enumerate only the cells inside the defined display area
                IEnumerator enumerator = visibleRange.GetEnumerator();
                Console.WriteLine("Cells within the defined DisplayRange:");
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current is Cell cell)
                    {
                        Console.WriteLine($"{cell.Name}: {cell.Value}");
                    }
                }

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DisplayRangeDemo.xlsx");

                // Ensure the directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

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
