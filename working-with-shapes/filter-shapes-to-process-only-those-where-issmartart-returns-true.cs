// Title: Identify and Convert SmartArt Shapes in Excel with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through each worksheet's shapes, processes only those where IsSmartArt is true, converts each SmartArt to a GroupShape via GetResultOfSmartArt, reports the shape ID, name and grouped‑shape count, and saves the updated file.
// Keywords: Aspose.Cells SmartArt filter | IsSmartArt property .NET | Convert SmartArt to GroupShape | Excel shape iteration Aspose | C# Aspose.Cells workbook processing | SmartArt extraction Excel | GroupShape count Aspose | US developers Aspose.Cells | UK .NET Excel automation | India C# Excel library
// Common Searches: Aspose.Cells loop through shapes and select SmartArt | Get GroupShape from SmartArt using Aspose.Cells C# | Check if a shape is SmartArt in an Excel file | How to count shapes inside a SmartArt diagram with Aspose | Save workbook after modifying SmartArt Aspose.Cells
// Developer Intent: Locate SmartArt objects in a workbook and transform them into GroupShape instances for analysis or further manipulation.
// Use Cases: Extract every SmartArt diagram, convert it to a GroupShape, and tally the individual elements for reporting. | Log the ID and name of each SmartArt shape before conversion, then persist the workbook with the new GroupShape data. | Create a reusable method that returns a collection of GroupShape objects derived from all SmartArt shapes on a specified worksheet.
// AI Prompts: Generate C# code with Aspose.Cells that replaces each SmartArt shape in a workbook with a PNG image. | Write a method that returns a List<GroupShape> obtained from all SmartArt shapes in a given worksheet. | Explain best practices for handling null results from GetResultOfSmartArt and how to log appropriate warnings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates through each worksheet's shapes, processes only those where IsSmartArt is true, converts each SmartArt to a GroupShape via GetResultOfSmartArt, reports the shape ID, name and grouped‑shape count, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        Console.WriteLine($"SmartArt found - Id: {shape.Id}, Name: {shape.Name}");

                        // Convert the SmartArt to a GroupShape (if possible) and inspect its contents
                        GroupShape groupShape = shape.GetResultOfSmartArt();
                        if (groupShape != null)
                        {
                            // Get the collection of grouped shapes and count them
                            int groupedCount = groupShape.GetGroupedShapes().Count();
                            Console.WriteLine($"Converted to GroupShape containing {groupedCount} shapes.");
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
