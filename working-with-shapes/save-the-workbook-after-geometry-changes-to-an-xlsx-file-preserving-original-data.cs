// Title: Save an XLSX workbook after changing column width, row height, and adding a rectangle shape while preserving original data using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing XLSX file with Aspose.Cells, sets column A width to 20 characters, row 1 height to 30 points, inserts a rectangle shape at row 5 column 5 with specific dimensions, and saves the workbook to a new XLSX file while keeping all original content unchanged. | Create a reusable C# method that receives a Workbook, column‑width, row‑height, and shape parameters, applies those geometry changes using Aspose.Cells, and writes the result as an XLSX file without losing any existing worksheet data.
// Common Searches: aspocells change column width and row height keep original data | c# add rectangle shape to existing workbook and save as xlsx using Aspose.Cells | preserve worksheet content after modifying shape placement with Aspose.Cells | save workbook as xlsx after geometry updates Aspose.Cells .NET
// Tags: set column width Aspose.Cells | set row height Aspose.Cells | add rectangle shape Aspose.Cells | freefloating shape placement Aspose.Cells | save workbook as xlsx preserving data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing XLSX file, adjusts column A width and row 1 height, inserts a free‑floating rectangle shape with defined size and position, then saves the workbook as a new XLSX file while preserving all original worksheet data.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook (preserves all original data)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Geometry changes -----

            // Change column width (e.g., column A)
            sheet.Cells.SetColumnWidth(0, 20); // width in characters

            // Change row height (e.g., row 1)
            sheet.Cells.SetRowHeight(0, 30); // height in points

            // Add a rectangle shape and set its position and size
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape rect = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,          // upper left row
                5,          // upper left column
                100,        // top offset (points)
                150,        // left offset (points)
                50,         // height (points)
                100);       // width (points)

            // Configure shape placement
            rect.Placement = PlacementType.FreeFloating;
            rect.Top = 100;   // position from top of the sheet in points
            rect.Left = 150;  // position from left of the sheet in points

            // ----- Save the workbook -----
            // Save as XLSX, preserving all original content and the geometry changes made above
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
