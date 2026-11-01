// Title: How to retrieve a shape's top‑left cell coordinates (row and column) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get the row and column of a shape's upper left corner | Retrieve shape UpperLeftRow and UpperLeftColumn values in a .NET Excel workbook | Align a picture to a specific cell range by reading its top‑left cell with Aspose.Cells | C# code to obtain shape position indices in an Excel worksheet using Aspose.Cells | How to read shape cell offsets (row, column) from an existing Excel file with Aspose.Cells
// Tags: Aspose.Cells shape position API | C# retrieve shape UpperLeftRow UpperLeftColumn | Excel shape alignment with cell range | Aspose.Cells read shape cell indices | shape top-left coordinates .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet, locates a shape by name, and reads its UpperLeftRow and UpperLeftColumn properties to obtain the absolute top‑left cell coordinates. It outputs the row and column indices, assumes zero pixel offsets, converts offsets to points, and shows how to use this information for precise shape alignment.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shape by name; adjust the name as needed
                Shape shape = worksheet.Shapes["MyShape"];
                if (shape == null)
                {
                    Console.WriteLine("Error: Shape \"MyShape\" was not found in the worksheet.");
                    return;
                }

                // Get shape position information (row/column indices are zero‑based)
                int topLeftRow = shape.UpperLeftRow;
                int topLeftColumn = shape.UpperLeftColumn;

                // Offsets are not available in older Aspose.Cells versions; set to zero
                int offsetRowPixels = 0;
                int offsetColumnPixels = 0;

                // Output the retrieved values
                Console.WriteLine($"Shape Top‑Left Cell: Row={topLeftRow}, Column={topLeftColumn}");
                Console.WriteLine($"Pixel Offsets: RowOffset={offsetRowPixels}px, ColumnOffset={offsetColumnPixels}px");

                // Convert pixel offsets to points (1 pixel ≈ 0.75 point at 96 DPI)
                double rowOffsetPoints = offsetRowPixels * 0.75;
                double columnOffsetPoints = offsetColumnPixels * 0.75;
                Console.WriteLine($"Offsets in Points: RowOffset={rowOffsetPoints}pt, ColumnOffset={columnOffsetPoints}pt");

                // (Optional) Save the workbook if any modifications were made
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
