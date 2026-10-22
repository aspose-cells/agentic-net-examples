// Title: How to retrieve the Shapes collection from the active worksheet in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# source that loads an Excel file with Aspose.Cells, accesses the selected sheet, and returns the sheet's Shapes collection. | Provide a C# loop that walks through each Shape in the active worksheet's Shapes collection and writes the shape's Name and Type to the console. | Show C# code that extracts only picture-type shapes from the active sheet's Shapes collection and prints their dimensions.
// Common Searches: Aspose.Cells C# retrieve shapes from the selected worksheet | list all shape names and types in an Excel file using Aspose.Cells | filter picture shapes on the active sheet with Aspose.Cells .NET | example code to iterate worksheet shapes Aspose.Cells C#
// Tags: Aspose.Cells active sheet shape enumeration | C# extract Excel shape objects | Aspose.Cells picture shape filter | list shape metadata .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, obtains the currently selected worksheet, accesses its Shapes collection, and iterates through each shape, outputting the shape's name and type while handling missing files and unexpected errors.
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

            // Get the currently selected (active) worksheet
            Worksheet selectedSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Retrieve the Shapes collection of the selected worksheet
            ShapeCollection shapes = selectedSheet.Shapes;

            // Example: iterate through shapes (placeholder for further processing)
            foreach (Shape shape in shapes)
            {
                // Process each shape as needed
                // For demonstration, output the shape name and type
                Console.WriteLine($"Shape name: {shape.Name}, Type: {shape.Type}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
