// Title: Add a TextBox shape anchored to a specific cell (e.g., B2) in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to place a TextBox shape at cell B2 with custom width, height, top and left offsets, set its text and border thickness, then save the workbook. | Create a reusable method that converts an Excel cell address to zero‑based row and column indexes and adds a TextBox shape to the given worksheet using Aspose.Cells, allowing optional fill color.
// Common Searches: Aspose.Cells C# how to anchor a textbox to cell B2 | C# add textbox shape to Excel worksheet with specific size using Aspose.Cells | convert Excel cell reference to row and column indexes Aspose.Cells | set textbox border weight Aspose.Cells example | save workbook with textbox shape Aspose.Cells .NET
// Tags: add textbox shape Aspose.Cells | anchor textbox to cell coordinates Aspose.Cells | cell address to row column index Aspose.Cells | textbox border styling Aspose.Cells | save workbook with shapes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, converts the target cell address (e.g., B2) to zero‑based row and column indexes, adds a TextBox shape anchored to that cell with specified width, height, and pixel offsets, sets its displayed text and border weight (optionally a fill color), and saves the file as WorkbookWithTextbox.xlsx.
class AddTextboxExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell for the textbox (e.g., "B2")
            string targetCell = "B2";

            // Convert cell name to zero‑based row and column indexes
            int row, column;
            CellsHelper.CellNameToIndex(targetCell, out row, out column);

            // Define textbox size and offsets (in pixels)
            int topOffset = 5;      // distance from the top edge of the cell
            int leftOffset = 5;     // distance from the left edge of the cell
            int textboxHeight = 50; // height of the textbox
            int textboxWidth = 150; // width of the textbox

            // Add a textbox shape anchored to the specified cell
            TextBox textbox = sheet.Shapes.AddTextBox(row, column, topOffset, leftOffset, textboxHeight, textboxWidth);

            // Set the displayed text
            textbox.Text = "Hello, Aspose.Cells!";

            // Format the textbox border
            textbox.Line.Weight = 1.0; // border thickness

            // Optional: set fill color if the API version supports it
            // Uncomment the following line if FillColor is available:
            // textbox.FillColor = System.Drawing.Color.LightYellow;

            // Save the workbook
            string outputPath = "WorkbookWithTextbox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
