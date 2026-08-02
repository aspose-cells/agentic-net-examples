// Title: C# batch update of shape text box margins in multiple Excel workbooks using Aspose.Cells
// Description: A C# console app that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, iterates every worksheet and shape, and sets the TextBoxOptions top, bottom, left and right margins to a 5‑point standard before saving the file to an output directory.
// Keywords: Aspose.Cells | C# | .NET | Excel shape margins | batch processing workbooks | TextBoxOptions | set shape text margins | multiple workbook automation | Excel API example | GitHub Aspose.Cells sample
// Common Searches: How to change text box margins for all shapes in Excel with Aspose.Cells | Batch update shape margins in multiple .xlsx files C# | Iterate through worksheets and shapes using Aspose.Cells | Set uniform text box padding for Excel dashboards programmatically | Aspose.Cells example for updating shape TextBoxOptions
// Developer Intent: Apply a predefined margin to every shape’s text box across a collection of Excel files in one automated run.
// Use Cases: Enforce corporate design standards on Excel reports before distribution | Prepare template workbooks with consistent shape padding for multiple teams | Automate compliance checks for dashboard visual consistency | Mass‑update legacy spreadsheets after a style guideline change
// AI Prompts: Write C# code that uses Aspose.Cells to set 5‑point margins on all shape text boxes in every worksheet of a workbook. | Refactor the batch loop to include logging of each modified shape name and add robust exception handling. | Create a reusable method that accepts custom top, bottom, left, and right margin values and applies them to all shapes in a given workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# console app that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, iterates every worksheet and shape, and sets the TextBoxOptions top, bottom, left and right margins to a 5‑point standard before saving the file to an output directory.
class BatchShapeMarginUpdater
{
    // Standard margin values in points
    const double StandardTopMargin = 5.0;
    const double StandardBottomMargin = 5.0;
    const double StandardLeftMargin = 5.0;
    const double StandardRightMargin = 5.0;

    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the updated workbooks will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string inputPath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Load the workbook (lifecycle rule: use constructor)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Apply standard text margins using TextBoxOptions
                    shape.TextBoxOptions.TopMarginPt = StandardTopMargin;
                    shape.TextBoxOptions.BottomMarginPt = StandardBottomMargin;
                    shape.TextBoxOptions.LeftMarginPt = StandardLeftMargin;
                    shape.TextBoxOptions.RightMarginPt = StandardRightMargin;
                }
            }

            // Build output file path
            string fileName = Path.GetFileName(inputPath);
            string outputPath = Path.Combine(outputFolder, fileName);

            // Save the modified workbook (lifecycle rule: use Save method)
            workbook.Save(outputPath);
        }

        Console.WriteLine("Batch processing completed.");
    }
}
