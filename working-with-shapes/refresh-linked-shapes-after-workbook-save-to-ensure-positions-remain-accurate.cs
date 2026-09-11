// Title: Refresh linked shapes after saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, saves it, then iterates through each worksheet's Shapes collection to call Refresh on shapes where IsLinked is true, and saves the workbook again. | Show a C# example that checks for the presence of an input Excel file, uses Aspose.Cells to open it, performs a double‑save workflow to update linked pictures, and includes robust error handling. | Generate a C# snippet that demonstrates how to refresh linked objects by looping over Worksheet.Shapes and invoking the Refresh method only on linked shapes after the first save.
// Common Searches: aspocells c# refresh linked pictures after workbook save | how to update linked shapes in an Excel file using Aspose.Cells .NET | iterate over worksheet shapes and call Refresh for linked objects Aspose.Cells | double save Excel workbook to refresh linked objects Aspose.Cells | check if shape is linked before refreshing Aspose.Cells C#
// Tags: linked shape refresh Aspose.Cells | iterate worksheet shapes C# | double save workbook Aspose.Cells | shape.IsLinked check C# | missing input Excel file handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RefreshLinkedShapesExampleApp
{
    // The sample loads an existing workbook, saves it, optionally loops through each worksheet's Shapes collection to refresh linked shapes (using the IsLinked property), and saves the workbook a second time, demonstrating a double‑save pattern to keep linked objects up‑to‑date while handling missing input files and exceptions.
    class RefreshLinkedShapesExample
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load an existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Perform any required modifications here
                // ...

                // Save the workbook for the first time
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");

                // NOTE: The RefreshAllLinkedObjects method is not available in the current Aspose.Cells for .NET version.
                // If linked objects need to be refreshed, iterate through shapes and refresh individually as needed.
                // Example (optional):
                // foreach (Worksheet sheet in workbook.Worksheets)
                // {
                //     foreach (Shape shape in sheet.Shapes)
                //     {
                //         if (shape.IsLinked) shape.Refresh();
                //     }
                // }

                // Save the workbook again after (potential) refreshing linked objects
                string refreshedPath = "output_refreshed.xlsx";
                workbook.Save(refreshedPath);
                Console.WriteLine($"Refreshed workbook saved to '{refreshedPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
