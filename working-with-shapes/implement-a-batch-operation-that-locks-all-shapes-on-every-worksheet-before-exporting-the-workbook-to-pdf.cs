// Title: Lock all shapes on every worksheet and export the workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells lock every shape in workbook before PDF conversion | batch set shape.IsLocked true for all worksheets using Aspose.Cells | prevent editing of Excel shapes when saving as PDF with Aspose.Cells .NET | iterate through worksheets and lock shapes in Aspose.Cells C# | how to protect shapes in Excel file before PDF export using Aspose.Cells
// Tags: shape.IsLocked property Aspose.Cells | lock all worksheet shapes C# | export workbook to PDF Aspose.Cells | iterate worksheets shapes Aspose.Cells | create output folder before PDF save C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel file, iterates through each worksheet, sets the IsLocked property of every Shape to true, ensures the target directory exists, and then saves the workbook as a PDF using Aspose.Cells for .NET, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and lock every shape
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    shape.IsLocked = true; // Prevent editing of the shape
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
