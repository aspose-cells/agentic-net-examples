// Title: Add a rectangle shape linking to Summary!A1 on each worksheet of an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a workbook, loops through every worksheet, inserts a rectangle shape at row 3 column C, sets its hyperlink to Summary!A1, and writes a log entry for each sheet. | Change the example to use an Oval shape instead of a rectangle and point the hyperlink to Summary!B2 while keeping the try‑catch logging structure. | Add a try‑catch around the workbook.Save call that records a custom success or failure message to the console.
// Common Searches: how to add a shape with hyperlink to each sheet using Aspose.Cells C# | Aspose.Cells iterate worksheets insert rectangle shape linking to summary cell | C# Aspose.Cells add hyperlink shape to Summary!A1 on all worksheets | log shape insertion results when adding shapes to Excel sheets with Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | shape hyperlink to Summary!A1 Aspose.Cells | iterate worksheets insert shape Aspose.Cells | error handling shape insertion Aspose.Cells | log shape addition results C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads 'input.xlsx', iterates each worksheet, adds a rectangle shape at row 3 column C with a hyperlink to Summary!A1, logs success or failure per sheet, and saves the modified workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the target summary cell (e.g., Summary!A1)
            string summarySheetName = "Summary";
            string summaryCellAddress = "A1";

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add a rectangle shape to the current worksheet
                    // Parameters: type, upper‑left row, upper‑left column, top offset, left offset, height, width
                    Shape shape = sheet.Shapes.AddShape(
                        MsoDrawingType.Rectangle, // shape type
                        2,    // upper‑left row (zero‑based, row 3)
                        2,    // upper‑left column (zero‑based, column C)
                        0,    // top offset in pixels
                        0,    // left offset in pixels
                        100,  // height in points
                        200   // width in points
                    );

                    // Configure the hyperlink to point to the summary cell
                    shape.Hyperlink.Address = $"{summarySheetName}!{summaryCellAddress}";
                    // Optional: set a screen tip
                    shape.Hyperlink.ScreenTip = $"Go to {summarySheetName}!{summaryCellAddress}";

                    // Optionally, give the shape a name or text
                    shape.Name = $"LinkToSummary_{sheet.Name}";
                    shape.Text = "Go to Summary";

                    // Log the operation
                    Console.WriteLine($"Added shape to worksheet '{sheet.Name}' linking to '{summarySheetName}!{summaryCellAddress}'.");
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Failed to add shape to worksheet '{sheet.Name}': {exShape.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
