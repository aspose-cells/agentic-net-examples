// Title: How to list the names of SmartArt shapes in every worksheet of an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook and prints each worksheet's SmartArt shape names. | Create a method that returns a Dictionary<string, List<string>> mapping worksheet names to their SmartArt shape names using Aspose.Cells. | Modify the example to also output the SmartArt layout type (e.g., Hierarchy, Process) for each identified shape.
// Common Searches: aspnet list smartart objects in an excel workbook using aspose.cells | c# enumerate smartart items in each worksheet of an xlsx file | how to retrieve smartart layout type with aspose.cells .net | get smartart shape properties from excel file using aspose.cells c# | iterate over worksheet shapes and filter smartart with aspose.cells
// Tags: aspocells enumerate smartart shapes | c# extract smartart identifiers | aspocells get smartart layout type | excel worksheet shape iteration .net | smartart shape detection using aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates through each worksheet and its Shapes collection, checks the IsSmartArt flag, and writes the worksheet name together with each SmartArt shape's Name (and optionally its layout type) to the console, while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Identify SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Output the SmartArt shape's name and its worksheet
                        Console.WriteLine($"Worksheet: {sheet.Name}, SmartArt Name: {shape.Name}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
