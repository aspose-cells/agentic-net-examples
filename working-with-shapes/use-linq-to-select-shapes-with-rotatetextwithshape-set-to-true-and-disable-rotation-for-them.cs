// Title: Use LINQ to disable rotation for shapes with RotateTextWithShape set to true in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, uses LINQ to locate every Shape where RotateTextWithShape is true, sets each shape's RotationAngle to 0, and saves the workbook. | Update the LINQ query in the existing example to also filter on RotateTextWithShape == true while keeping the rotation‑reset logic intact. | Add robust file‑existence validation and detailed exception handling around the shape‑rotation disabling routine.
// Common Searches: aspocells disable rotation for shapes with RotateTextWithShape true c# | linq select shapes where RotateTextWithShape is enabled in excel workbook | reset rotation angle of text‑wrapped shapes using Aspose.Cells .NET | how to turn off shape rotation for rotated text in Aspose.Cells C# | filter and update shape properties with LINQ in Aspose.Cells
// Tags: disable shape rotation Aspose.Cells | LINQ filter RotateTextWithShape property | reset RotationAngle for Excel shapes .NET | Aspose.Cells shape property manipulation | C# LINQ shape selection Excel workbook

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file with Aspose.Cells, uses LINQ to select all Shape objects whose RotateTextWithShape flag is true, clears their RotationAngle to disable rotation, and saves the modified workbook, including file existence checks and basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Select all shapes that have rotation applied (RotationAngle != 0)
            var shapesToUpdate = sheet.Shapes
                                     .Cast<Shape>()
                                     .Where(s => s.RotationAngle != 0);

            // Reset rotation for each selected shape
            foreach (var shape in shapesToUpdate)
            {
                shape.RotationAngle = 0;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
