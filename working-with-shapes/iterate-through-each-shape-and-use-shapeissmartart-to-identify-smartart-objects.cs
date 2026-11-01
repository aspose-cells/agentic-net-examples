// Title: Iterate through worksheet shapes and detect SmartArt objects using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet, examines each Shape, and prints the shape name together with its upper‑left row and column when the shape is SmartArt. | Modify the example to gather SmartArt shape information into a collection and write the details (sheet name, shape name, row, column) to a CSV file.
// Common Searches: aspocells c# find smartart shapes in excel workbook | how to list smartart objects and their cell positions with Aspose.Cells .NET | enumerate worksheet shapes and check IsSmartArt property using Aspose.Cells
// Tags: detect smartart shapes aspocells | enumerate worksheet shape collection c# | shape issmartart property aspocells | smartart position upperleftrow aspocells | aspocells shape collection example

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads a workbook, iterates each worksheet's ShapeCollection, checks Shape.IsSmartArt, and outputs the SmartArt shape name with its upper‑left row and column; optionally saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Iterate through each shape
                foreach (Shape shape in shapes)
                {
                    // Identify SmartArt objects
                    if (shape.IsSmartArt)
                    {
                        // Use UpperLeftRow/UpperLeftColumn to obtain shape position
                        int row = shape.UpperLeftRow;
                        int column = shape.UpperLeftColumn;
                        Console.WriteLine($"SmartArt found in sheet '{sheet.Name}': Shape Name = {shape.Name}, Row = {row}, Column = {column}");
                    }
                }
            }

            // Save the workbook (optional, if modifications were made)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
