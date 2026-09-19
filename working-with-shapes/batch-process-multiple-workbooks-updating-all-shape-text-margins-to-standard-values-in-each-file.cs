// Title: Batch update shape text alignment (and prepare for margin settings) across multiple Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Generate a C# console application that scans a folder for .xlsx files, opens each workbook with Aspose.Cells, iterates every worksheet and shape, sets TextHorizontalAlignment and TextVerticalAlignment to Center for shapes containing text, and saves the modified files to a separate output directory. | Extend the batch shape processor to log the file name and any exception messages to a text file while continuing processing of remaining workbooks. | Modify the code to include placeholders for future shape text margin properties, showing where StandardLeftMargin, StandardRightMargin, StandardTopMargin, and StandardBottomMargin would be applied when Aspose.Cells adds margin support. | Add a command‑line argument that lets the user specify the desired horizontal and vertical alignment values (e.g., Center, Left, Top) for shape text during batch processing.
// Common Searches: asp.net c# batch process excel workbooks to center shape text using Aspose.Cells | how to iterate through all shapes in every worksheet of multiple Excel files with Aspose.Cells | automate updating shape text alignment across a folder of .xlsx files in C# | Aspose.Cells set shape text horizontal alignment for all worksheets programmatically | C# script to apply consistent shape formatting to many Excel workbooks
// Tags: batch shape text alignment Aspose.Cells | iterate worksheets shapes c# | process multiple excel workbooks Aspose.Cells | center shape text Aspose.Cells | prepare shape margin placeholders Aspose.Cells | error logging batch excel processing c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads every .xlsx file from a specified input directory, walks through each worksheet and each shape, centers the horizontal and vertical text alignment for shapes that contain text, and saves the updated workbooks to an output folder. It also demonstrates where margin constants could be applied in future API versions and includes basic error handling for robust batch processing.
class ShapeMarginBatchProcessor
{
    // Standard margin values (in points) – retained for reference; Aspose.Cells Shape does not expose margin properties.
    const double StandardLeftMargin = 5.0;
    const double StandardRightMargin = 5.0;
    const double StandardTopMargin = 5.0;
    const double StandardBottomMargin = 5.0;

    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\Workbooks\Input";
        // Folder where the updated workbooks will be saved
        string outputFolder = @"C:\Workbooks\Output";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Process each Excel file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only shapes that contain text
                        if (!string.IsNullOrEmpty(shape.Text))
                        {
                            // Optional alignment settings
                            shape.TextHorizontalAlignment = TextAlignmentType.Center;
                            shape.TextVerticalAlignment = TextAlignmentType.Center;

                            // Aspose.Cells does not provide direct margin properties for Shape text.
                            // If future versions add such properties, they can be set here using the
                            // Standard*Margin constants defined above.
                        }
                    }
                }

                // Determine output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the updated workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
