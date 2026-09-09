// Title: Add uniquely named TextBox shapes to every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, iterates all worksheets, inserts a TextBox shape at a given cell location, and assigns a distinct Name to each TextBox. | Adjust the TextBox insertion logic to compute row, column, and pixel offsets based on each worksheet's layout while ensuring sequential identifiers are applied without overwriting existing shapes.
// Common Searches: aspnet c# add textbox shape to each worksheet aspose.cells example | assign incremental names to shapes when looping through worksheets in Aspose.Cells | batch create text boxes in all sheets of an Excel workbook using Aspose.Cells for .NET | set textbox size and position programmatically with Aspose.Cells C#
// Tags: bulk shape insertion aspose.cells | sequential shape identifiers .net | worksheet loop shape creation c# | textbox geometry parameters aspose.cells | excel workbook shape handling .net

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel file, loops through every worksheet, adds a TextBox shape at a fixed cell with defined pixel dimensions, sets its displayed text, assigns a unique Name using an incrementing counter, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Paths to the source and destination Excel files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Counter to generate unique identifiers for each TextBox
        int textboxCounter = 0;

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Define position and size for the TextBox (row, column, top, left, height, width)
            int upperLeftRow = 1;          // Row index (0‑based)
            int upperLeftColumn = 1;       // Column index (0‑based)
            int top = 0;                   // Pixels from the top of the cell
            int left = 0;                  // Pixels from the left of the cell
            int height = 100;              // Height in pixels
            int width = 200;               // Width in pixels

            // Add a TextBox shape to the current worksheet
            TextBox textBox = sheet.Shapes.AddTextBox(
                upperLeftRow, upperLeftColumn, top, left, height, width);

            // Set the displayed text (optional)
            textBox.Text = $"TextBox on sheet \"{sheet.Name}\"";

            // Assign a unique identifier using the Name property
            textBox.Name = $"TextBox_{textboxCounter}";

            // Increment the counter for the next TextBox
            textboxCounter++;
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save(outputPath);
    }
}
