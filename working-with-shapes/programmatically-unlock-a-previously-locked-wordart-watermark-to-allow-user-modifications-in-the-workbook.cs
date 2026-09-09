// Title: Unlock a locked WordArt watermark in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Iterate through every worksheet's ShapeCollection with Aspose.Cells and set Shape.IsLocked = false, then save the workbook. | Programmatically remove the lock from WordArt objects in an existing .xlsx file using C# and Aspose.Cells. | Load an Excel file, unlock all shapes including WordArt watermarks, and write the updated file to a new location.
// Common Searches: aspnet unlock WordArt watermark in existing Excel file using Aspose.Cells | c# Aspose.Cells set shape IsLocked false for all shapes in workbook | how to edit locked WordArt objects in .xlsx with Aspose.Cells | remove shape lock from Excel watermark programmatically .NET
// Tags: shape.IsLocked property Aspose.Cells | modify WordArt lock status Excel | iterate worksheet ShapeCollection C# | save workbook after unlocking shapes Aspose.Cells | programmatic Excel watermark editing .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing .xlsx file, loops through each worksheet's ShapeCollection, sets each Shape's IsLocked property to false (unlocking WordArt watermarks), and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains the locked WordArt watermark
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of shapes (including WordArt) on the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Unlock each shape; this includes WordArt objects
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];
                    shape.IsLocked = false;
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the unlocked WordArt watermark
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
