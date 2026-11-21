// Title: Replace SmartArt and other shape text across all worksheets of an XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace SmartArt text in Excel workbook | update shape text for all worksheets using Aspose.Cells .NET | iterate through Excel shapes and set Text property with Aspose.Cells | C# change SmartArt node text in XLSX file using Aspose.Cells | save modified Excel workbook after editing shape text with Aspose.Cells
// Tags: replace SmartArt shape text Aspose.Cells | modify shape.Text across worksheets C# | bulk update Excel shape text .NET | Aspose.Cells workbook save after shape edit | handle shape processing exceptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an XLSX file, loops through every worksheet and each shape, sets the Text property to a new value for SmartArt and regular text shapes, catches shape‑level errors, ensures the output directory exists, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // If the shape is a SmartArt, update its text (fallback to shape text if nodes are unavailable)
                        if (shape.IsSmartArt)
                        {
                            // Aspose.Cells may not expose SmartArt nodes directly in all versions.
                            // As a safe alternative, set the shape's text which updates the displayed SmartArt text.
                            shape.Text = "Updated Text";
                        }
                        else if (!string.IsNullOrEmpty(shape.Text))
                        {
                            // For regular text-containing shapes, update the text as well
                            shape.Text = "Updated Text";
                        }
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Error processing shape on sheet \"{sheet.Name}\": {shapeEx.Message}");
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
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
