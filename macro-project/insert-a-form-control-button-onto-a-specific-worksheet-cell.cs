// Title: Insert a Form Control button into cell B3 of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Insert a Form Control button at row 2, column 1 (cell B3), set its height to 30 px, width to 100 px, label it "Click Me", and save the workbook as FormControlButton.xlsx using Aspose.Cells for .NET. | Programmatically create a new workbook, ensure the target directory exists, add a button shape with custom offsets, assign text, and write the file to disk with the Aspose.Cells C# API.
// Common Searches: Aspose.Cells how to add a form control button to cell B3 in C# | C# Aspose.Cells place a button at a specific row and column | Setting button text and size when inserting a form control with Aspose.Cells | Saving an Excel file that contains a button using Aspose.Cells .NET | Creating a new workbook and adding a clickable button programmatically with Aspose.Cells
// Tags: aspocells button shape insertion | aspocells button placement by cell coordinates | aspocells set button caption and dimensions | aspocells workbook save with controls | aspocells ensure output folder exists

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first Worksheet, defines cell B3 (row index 2, column index 1), adds a Form Control button shape with a height of 30 px and width of 100 px, sets the caption to "Click Me", ensures the output directory exists, saves the file as FormControlButton.xlsx, and handles any exceptions.
class InsertFormControlButton
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell (e.g., B3 -> row index 2, column index 1)
            int targetRow = 2;      // zero‑based row index
            int targetColumn = 1;   // zero‑based column index

            // Add a button shape to the worksheet.
            // Parameters: upper‑left row, upper‑left column,
            // top offset, left offset, height (pixels), width (pixels)
            Button button = sheet.Shapes.AddButton(
                targetRow, targetColumn,
                0, 0,
                30,   // height
                100); // width

            // Set the button caption
            button.Text = "Click Me";

            // Define output file path
            string outputPath = "FormControlButton.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
