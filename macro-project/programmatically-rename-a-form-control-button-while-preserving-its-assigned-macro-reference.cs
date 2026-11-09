// Title: Rename an Excel form control button while preserving its macro using Aspose.Cells for .NET
// Common Searches: C# code to rename a form control button in an Excel file without breaking its macro link using Aspose.Cells | How to change the name of a button shape in a workbook while keeping the VBA macro reference | Aspose.Cells rename button shape programmatically and retain macro association | Update Excel form control button name via Aspose.Cells .NET API | Preserve macro assignment when renaming an Excel button with Aspose.Cells
// Tags: Aspose.Cells rename button shape | form control button macro preservation | C# modify Excel button name | shape name change Aspose.Cells | retain VBA macro link after button rename

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, locates a form control button named "Button 1" on the first worksheet, renames it to "NewButtonName" while automatically retaining its associated VBA macro, and saves the updated file.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the form control button by its current name
            Button buttonShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                // Identify the button by type and name
                if (shape is Button btn && shape.Name == "Button 1")
                {
                    buttonShape = btn;
                    break;
                }
            }

            if (buttonShape != null)
            {
                // Rename the button; macro association is retained automatically
                buttonShape.Name = "NewButtonName";
            }
            else
            {
                Console.WriteLine("Button with name 'Button 1' not found.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Basic error handling
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
