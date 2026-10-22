// Title: Load an Excel workbook, select a worksheet, and add a free‑floating rectangle shape with text using Aspose.Cells for .NET
// AI Prompts: Read an existing Excel file, choose a worksheet by name or index, insert a rectangle shape at row 2 column 1 with width 50 and height 100, set its placement to FreeFloating, assign the text "Sample Shape", and save the result to a new file. | Before opening the workbook, verify that the source Excel file exists, handle any exceptions that may occur, and ensure the modified workbook is saved after the shape is added.
// Common Searches: Aspose.Cells C# how to add a rectangle shape to a specific worksheet after loading the workbook | C# code to verify Excel file exists before using Aspose.Cells to modify it | Insert free‑floating shape with text into Excel using Aspose.Cells .NET API | Select worksheet by name in Aspose.Cells before drawing shapes
// Tags: load workbook select worksheet Aspose.Cells | shape creation Aspose.Cells | freefloating placement Aspose.Cells | excel file existence check C# | save workbook after modifications Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads 'input.xlsx', selects 'Sheet1', adds a free‑floating rectangle shape containing the text "Sample Shape", and saves the updated file as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Select the target worksheet (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape rectangle = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // correct enum for shape type
                2, 1, 0, 0, 100, 50);
            rectangle.Placement = PlacementType.FreeFloating;
            rectangle.Text = "Sample Shape";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
