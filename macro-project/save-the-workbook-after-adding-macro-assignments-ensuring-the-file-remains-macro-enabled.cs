// Title: Insert a rectangular button shape, assign a macro, and save the workbook as a macro‑enabled XLSM file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a free‑floating rectangle shape to the first worksheet, sets its Name and Caption, links it to a macro identifier, and saves the workbook as an XLSM file while preserving existing macros using Aspose.Cells. | Demonstrate how to verify the source .xlsm file, create the destination folder if it does not exist, and call Workbook.Save with SaveFormat.Xlsm to keep macros intact. | Explain the steps required to enable macro assignment on a shape after upgrading to a version of Aspose.Cells that supports the Shape.Macro property.
// Common Searches: how to add a button shape to an Excel worksheet with Aspose.Cells C# | assign a macro to a shape in a macro‑enabled workbook using Aspose.Cells | save an Aspose.Cells workbook as .xlsm while preserving macros | ensure output directory exists before calling Workbook.Save in C# | Aspose.Cells shape macro property not available older version
// Tags: add rectangle shape button Aspose.Cells | macro assignment to shape Aspose.Cells | save workbook as xlsm preserving macros | create output directory before saving Aspose.Cells | macro‑enabled workbook handling Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// // Loads a macro‑enabled workbook, inserts a free‑floating rectangular shape named 'MyButton' with the caption 'Run Macro', optionally links it to a macro, ensures the output folder exists, and saves the file as an .xlsm preserving all macros.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.xlsm";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the macro‑enabled workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangular shape (used as a placeholder for a button)
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape button = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // generic rectangle shape
                1, 0, 0, 100, 30, 80);    // row, column, top, left, height, width

            button.Name = "MyButton";
            button.Placement = PlacementType.FreeFloating;
            button.Text = "Run Macro";

            // Note: Assigning a macro to a shape requires a newer Aspose.Cells version.
            // If supported, you can uncomment the following line after upgrading:
            // button.Macro = "MyMacro";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook preserving macros
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
