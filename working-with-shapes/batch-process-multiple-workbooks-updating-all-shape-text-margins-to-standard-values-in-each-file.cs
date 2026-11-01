// Title: Batch update shape text alignment margins in multiple Excel workbooks with Aspose.Cells for .NET
// Common Searches: c# aspocells batch modify shape text horizontal and vertical alignment in a folder of .xlsx files | how to iterate all worksheets and shapes to set text margins using Aspose.Cells .NET | process multiple Excel workbooks to center shape text using Aspose.Cells C# | automate shape text alignment updates across many spreadsheets with Aspose.Cells | C# script to apply standard text margins to all shapes in Excel files
// Tags: batch update shape text alignment Aspose.Cells | set shape TextHorizontalAlignment TextVerticalAlignment C# | iterate worksheets shapes Aspose.Cells | process multiple .xlsx workbooks Aspose.Cells | center shape text margins Excel C# | automate shape text formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The C# program scans a specified directory for .xlsx files, loads each workbook with Aspose.Cells, iterates through every worksheet and its shapes, and for shapes containing text sets TextHorizontalAlignment and TextVerticalAlignment to Center, then saves the workbook, effectively standardizing shape text margins across all workbooks in the folder.
class BatchShapeMarginUpdater
{
    // Standard margin values can be represented by centering the text within the shape.
    // Adjust these values if different margin settings are required.
    private static readonly TextAlignmentType StandardHorizontalAlignment = TextAlignmentType.Center;
    private static readonly TextAlignmentType StandardVerticalAlignment = TextAlignmentType.Center;

    static void Main()
    {
        // Folder containing the workbooks to process.
        string folderPath = @"C:\Workbooks";

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all Excel files in the folder (you can adjust the search pattern as needed).
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in workbookFiles)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet.
                    foreach (Shape shape in sheet.Shapes)
                    {
                        try
                        {
                            // Process only shapes that contain text.
                            if (!string.IsNullOrEmpty(shape.Text))
                            {
                                // Update text alignment to represent standard margins.
                                shape.TextHorizontalAlignment = StandardHorizontalAlignment;
                                shape.TextVerticalAlignment = StandardVerticalAlignment;
                            }
                        }
                        catch (Exception shapeEx)
                        {
                            Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {shapeEx.Message}");
                        }
                    }
                }

                // Save the workbook, overwriting the original file.
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("All shape text margins have been updated.");
    }
}
