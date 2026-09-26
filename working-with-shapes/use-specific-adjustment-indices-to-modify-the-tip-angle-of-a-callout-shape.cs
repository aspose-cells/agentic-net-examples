// Title: Change the tip angle of a callout shape in an existing Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file, retrieves the first shape, and sets its first adjustment value to alter the callout tip angle using Aspose.Cells. | Show how to use reflection in C# to access the Shape.Adjustments property in Aspose.Cells, modify the adjustment array, and save the workbook.
// Common Searches: asp.net how to adjust callout shape tip angle in Excel using Aspose.Cells | c# set shape adjustments array for callout in existing workbook | modify callout shape parameters with Aspose.Cells reflection example | change callout arrow angle programmatically in .xlsx via Aspose.Cells | update shape adjustment index in Excel file using Aspose.Cells for .NET
// Tags: adjust callout shape tip angle Aspose.Cells | shape adjustments array reflection C# | modify shape parameters Excel workbook .NET | set shape adjustment index Aspose.Cells | callout shape adjustment property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing .xlsx file, checks for shapes on the first worksheet, uses reflection to obtain the Shape.Adjustments array, sets the first adjustment value (e.g., 0.5) to change the callout tip angle, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
            }
            else
            {
                // Access the first shape (could be any type)
                Shape shape = sheet.Shapes[0];

                // Attempt to modify adjustments if the property is available
                try
                {
                    // Some shape types (e.g., Callout) support adjustments.
                    // Use reflection to avoid compile‑time dependency on the Adjustments property.
                    var adjustmentsProp = typeof(Shape).GetProperty("Adjustments");
                    if (adjustmentsProp != null)
                    {
                        var adjustments = adjustmentsProp.GetValue(shape) as double[];
                        if (adjustments != null && adjustments.Length > 0)
                        {
                            adjustments[0] = 0.5; // Example adjustment
                            Console.WriteLine("Shape adjustment applied.");
                        }
                        else
                        {
                            Console.WriteLine("Shape does not support adjustments or has no adjustable parameters.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Adjustments property not found on Shape.");
                    }
                }
                catch (Exception adjEx)
                {
                    Console.WriteLine($"Adjustment error: {adjEx.Message}");
                }
            }

            // Ensure the output directory exists
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
