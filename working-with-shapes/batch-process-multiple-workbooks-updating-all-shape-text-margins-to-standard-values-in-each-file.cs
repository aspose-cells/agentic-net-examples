// Title: Batch update shape text box margins in multiple Excel workbooks with Aspose.Cells (C#)
// Description: Scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, iterates every worksheet and shape, sets left, right, top, and bottom text box margins to a standard 5‑point value, optionally fits the shape to the new margins, and saves the modified files to a separate output directory.
// Keywords: Aspose.Cells | C# | Excel shape margins | text box margins | batch processing | multiple workbooks | standard margin | FitToTextSize | automation | office document API
// Common Searches: How to set uniform text box margins for all shapes in many Excel files using Aspose.Cells | C# batch update shape margins across a folder of workbooks | Aspose.Cells example for changing shape text box margins in bulk | Automate margin standardization for Excel shapes with .NET
// Developer Intent: Apply a consistent text box margin to every shape in each workbook of a batch of Excel files.
// Use Cases: Enforce corporate branding by standardizing shape margins in all generated reports. | Prepare a collection of spreadsheets for printing, ensuring uniform text box spacing. | Automate cleanup of legacy workbooks that contain inconsistent shape formatting.
// AI Prompts: Write C# code with Aspose.Cells that sets all shape text box margins to 8 points in a single workbook. | Refactor the batch margin updater to add per‑file logging and skip shapes lacking TextBoxOptions without throwing errors. | Explain how to extend the processor to also change each shape's fill color while updating its margins.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, iterates every worksheet and shape, sets left, right, top, and bottom text box margins to a standard 5‑point value, optionally fits the shape to the new margins, and saves the modified files to a separate output directory.
class BatchShapeMarginUpdater
{
    // Standard margin value in points
    const double StandardMargin = 5.0;

    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the updated workbooks will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder (you can adjust the pattern as needed)
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string inputPath in workbookFiles)
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Only process shapes that support text box options (e.g., text boxes)
                    // Some shapes may not have TextBoxOptions; skip those.
                    if (shape.TextBoxOptions != null)
                    {
                        // Set all four margins to the standard value
                        shape.TextBoxOptions.LeftMarginPt = StandardMargin;
                        shape.TextBoxOptions.RightMarginPt = StandardMargin;
                        shape.TextBoxOptions.TopMarginPt = StandardMargin;
                        shape.TextBoxOptions.BottomMarginPt = StandardMargin;

                        // Optionally, recalculate the shape size to fit the new margins
                        shape.FitToTextSize();
                    }
                }
            }

            // Determine output file path (same name, different folder)
            string fileName = Path.GetFileName(inputPath);
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the updated workbook
            workbook.Save(outputPath);
        }

        Console.WriteLine("Batch processing completed.");
    }
}
