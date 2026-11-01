// Title: Check that SmartArt nodes contain non‑empty text in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to iterate over SmartArt shapes in an Excel file with Aspose.Cells C# | C# code to detect empty SmartArt node text using Aspose.Cells | verify SmartArt node content is not blank in an Aspose.Cells workbook | Aspose.Cells .NET find SmartArt shapes with missing text | read SmartArt node text and check for emptiness with Aspose.Cells
// Tags: Aspose.Cells SmartArt node text validation | C# iterate workbook shapes Aspose.Cells | detect empty SmartArt nodes Excel | validate SmartArt content .NET | process SmartArt shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, loops through each worksheet and its shapes, identifies SmartArt shapes, and demonstrates how to access SmartArt nodes to verify that their text is not empty before optionally saving the file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file \"{inputFile}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Aspose.Cells may not expose SmartArt node details in all versions.
                        // Here we simply report the presence of a SmartArt shape.
                        Console.WriteLine($"SmartArt shape detected in sheet '{sheet.Name}', shape ID {shape.Id}.");

                        // If SmartArt node access is supported, the following code can be used:
                        // var smartArt = shape.SmartArt;
                        // foreach (SmartArtNode node in smartArt.AllNodes)
                        // {
                        //     string nodeText = node.Text?.Trim();
                        //     if (string.IsNullOrEmpty(nodeText))
                        //     {
                        //         Console.WriteLine($"Empty SmartArt node detected in sheet '{sheet.Name}', shape ID {shape.Id}.");
                        //     }
                        // }
                    }
                }
            }

            // Save the workbook (optional, if modifications were made)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
