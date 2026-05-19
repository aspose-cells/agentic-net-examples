using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class BatchUpdateShapeMargins
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the updated workbooks will be saved
        string outputFolder = @"C:\UpdatedWorkbooks";

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        // Standard margin values (in points)
        const double leftMargin = 5.0;
        const double rightMargin = 5.0;
        const double topMargin = 5.0;
        const double bottomMargin = 5.0;

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(filePath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Only shapes that support text box options have a non‑null TextBoxOptions property
                    if (shape.TextBoxOptions != null)
                    {
                        // Update the text margins to the standard values
                        shape.TextBoxOptions.LeftMarginPt = leftMargin;
                        shape.TextBoxOptions.RightMarginPt = rightMargin;
                        shape.TextBoxOptions.TopMarginPt = topMargin;
                        shape.TextBoxOptions.BottomMarginPt = bottomMargin;
                    }
                }
            }

            // Save the modified workbook to the output folder (uses the provided save rule)
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            workbook.Save(outputPath);

            // Release resources
            workbook.Dispose();
        }

        Console.WriteLine("Batch processing completed.");
    }
}