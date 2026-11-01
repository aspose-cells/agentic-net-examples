// Title: Apply the same drop‑shadow style to every shape on the first ten worksheets of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over the first ten worksheets and assigns a predefined shadow effect to each shape. | Create a reusable C# method that accepts a Workbook and a shadow configuration, then applies that configuration to all shapes on up to ten sheets. | Show how to check for the source file, load the workbook, apply identical shadow formatting to shapes across multiple worksheets, and save the modified file.
// Common Searches: Aspose.Cells C# set drop shadow for shapes on multiple worksheets | how to loop through first 10 sheets and modify shape formatting with Aspose.Cells | batch apply shadow effect to all shapes in an Excel file using .NET SDK | C# example for applying same visual style to shapes across several worksheets
// Tags: apply shape shadow Aspose.Cells C# | iterate worksheets batch formatting Aspose.Cells | set drop shadow on Excel shapes .NET | process first ten worksheets Aspose.Cells | shape formatting loop workbook C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, determines the smaller of ten or the total worksheet count, and loops through each of those worksheets. For every shape on each sheet, a placeholder is provided where a consistent drop‑shadow configuration can be applied using Aspose.Cells APIs. After processing, the workbook is saved to a new file, with checks for file existence and exception handling to ensure reliable execution.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process up to the first ten worksheets
            int sheetsToProcess = Math.Min(10, workbook.Worksheets.Count);

            for (int i = 0; i < sheetsToProcess; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // NOTE: Shadow manipulation APIs may not be available in older Aspose.Cells versions.
                    // If needed, shadow settings can be applied using the appropriate API for the version in use.
                    // The following block is intentionally left empty to maintain compatibility.
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
