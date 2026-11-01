// Title: Replace a <TAG_1> placeholder in the HeaderBox TextBox shape of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace placeholder text in a specific TextBox shape | How to find and edit a shape by its name in an Excel file with Aspose.Cells | Programmatically update TextBox content in a .xlsx using Aspose.Cells .NET | Replace <TAG_1> with a dynamic title in an Excel header shape using C# | Iterate through worksheets to modify named shape text with Aspose.Cells
// Tags: Aspose.Cells replace TextBox shape text | C# retrieve shape using Name property Aspose.Cells | modify Excel header TextBox content Aspose.Cells | dynamic placeholder substitution in .xlsx via Aspose.Cells | process all worksheets for shape updates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, scans every worksheet to locate the TextBox shape named "HeaderBox", replaces the placeholder "<TAG_1>" with a dynamic title string, and saves the updated file to a new location.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string dynamicTitle = "Quarterly Sales Report";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets to locate the shape named "HeaderBox"
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Process only the shape with the target name
                        if (shape.Name == "HeaderBox")
                        {
                            // Replace the placeholder <TAG_1> with the dynamic title
                            shape.Text = shape.Text.Replace("<TAG_1>", dynamicTitle);
                        }
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Warning: Unable to process shape \"{shape.Name}\" on sheet \"{sheet.Name}\": {innerEx.Message}");
                    }
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
