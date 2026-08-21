// Title: Add a rectangle shape with custom size and anchor it to a cell range using Aspose.Cells for .NET
// Description: Creates a new workbook, inserts a rectangle shape with a 200 × 120 pixel size, and anchors it to the cell range B3:D5 (rows 2‑4, columns 1‑3) using the MoveToRange method before saving the file as XLSX.
// Keywords: Aspose.Cells add rectangle shape | custom shape size Aspose.Cells .NET | MoveToRange anchor shape | shape dimensions pixels Aspose.Cells | C# Aspose.Cells shape example | anchor shape to cell range
// Common Searches: Aspose.Cells add rectangle with specific width and height | How to anchor a shape to a range of cells in Aspose.Cells | MoveToRange method usage for shapes Aspose.Cells | Set shape size in pixels using Aspose.Cells C# | Aspose.Cells shape positioning example
// Developer Intent: Insert a rectangle of defined pixel dimensions and attach it to a chosen cell range in an Excel workbook programmatically.
// Use Cases: Overlay a colored box on a summary section (B3:D5) in an automated report. | Place a company logo sized 200 × 120 px and lock it to cells B2:D4 in a template. | Create a placeholder shape for user input that moves with cells C5:E7 when the sheet is edited.
// AI Prompts: Generate C# code with Aspose.Cells that adds an ellipse of 150 × 100 pixels and anchors it to cells C2:E4. | Explain how MoveToRange translates row/column indices and pixel offsets into shape positioning. | Write a reusable method that receives width, height, and a cell range, then returns a rectangle shape anchored to that range using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Creates a new workbook, inserts a rectangle shape with a 200 × 120 pixel size, and anchors it to the cell range B3:D5 (rows 2‑4, columns 1‑3) using the MoveToRange method before saving the file as XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the shape collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Define custom dimensions (in pixels)
                int customHeight = 120; // height of the shape
                int customWidth = 200;  // width of the shape

                // Add a rectangle shape with temporary position (0,0) and the custom size
                RectangleShape rectangle = shapes.AddRectangle(
                    topRow: 0,
                    top: 0,
                    leftColumn: 0,
                    left: 0,
                    height: customHeight,
                    width: customWidth);

                // Define the cell range to which the shape should be anchored
                int topRow = 2;    // upper‑left row index (zero‑based)
                int leftCol = 1;   // upper‑left column index (zero‑based)
                int bottomRow = 4; // lower‑right row index
                int rightCol = 3;  // lower‑right column index

                // Anchor the shape to the specified cell range
                rectangle.MoveToRange(topRow, leftCol, bottomRow, rightCol);

                // Output file path
                string outputPath = "ShapeAnchoredToRange.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
