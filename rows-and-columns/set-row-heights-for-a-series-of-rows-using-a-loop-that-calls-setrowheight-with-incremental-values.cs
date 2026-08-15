// Title: Set Incremental Row Heights in Excel with Aspose.Cells for .NET (C# Loop)
// Description: This example creates a new Workbook, accesses the first worksheet, and uses a C# for‑loop to assign each of the first ten rows a distinct height, starting at 15 pt and increasing by 2.5 pt per row via Cells.SetRowHeight. It also writes a label in column A showing the applied height and saves the file as RowHeightsLoopDemo.xlsx.
// Keywords: Aspose.Cells | SetRowHeight | C# Excel row height | incremental row height | loop row height | programmatic row sizing | Excel automation .NET | Cells.SetRowHeight example
// Common Searches: Aspose.Cells set row height in a loop | C# incremental row height Aspose.Cells | how to change multiple row heights programmatically Excel .NET | SetRowHeight example with varying heights | apply different heights to rows using Aspose.Cells
// Developer Intent: Programmatically apply different heights to multiple rows in a worksheet using a loop.
// Use Cases: Create a report where each successive row is taller to visually separate sections. | Generate a spreadsheet with progressively larger header rows for emphasis. | Add diagnostic text that displays each row's height to verify custom sizing in generated files.
// AI Prompts: Show how to adjust the loop so row height is calculated from the length of cell content with Aspose.Cells. | Provide a sample that reads a list of height values from an array and applies them to rows using Cells.SetRowHeight. | Explain how to revert rows back to the default height after custom heights have been set in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a new Workbook, accesses the first worksheet, and uses a C# for‑loop to assign each of the first ten rows a distinct height, starting at 15 pt and increasing by 2.5 pt per row via Cells.SetRowHeight. It also writes a label in column A showing the applied height and saves the file as RowHeightsLoopDemo.xlsx.
    public class SetRowHeightsWithLoop
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the number of rows to modify, the starting height, and the increment per row
                int totalRows = 10;
                double startHeight = 15.0;   // height for the first row (in points)
                double increment = 2.5;      // increase each subsequent row by this amount

                // Loop through the rows and set their heights using SetRowHeight
                for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
                {
                    double height = startHeight + (rowIndex * increment);
                    cells.SetRowHeight(rowIndex, height);
                }

                // Optionally, add some data to visualize the row heights
                for (int i = 0; i < totalRows; i++)
                {
                    cells[i, 0].PutValue($"Row {i + 1} height set to {cells.GetRowHeight(i)} points");
                }

                // Save the workbook to a file
                string outputPath = "RowHeightsLoopDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRowHeightsWithLoop.Run();
        }
    }
}
