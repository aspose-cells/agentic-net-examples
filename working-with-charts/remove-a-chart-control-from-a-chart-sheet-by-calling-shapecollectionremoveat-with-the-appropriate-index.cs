// Title: Delete a chart control from a chart sheet with Aspose.Cells ShapeCollection.RemoveAt in C#
// AI Prompts: Write C# code that opens an Excel workbook, locates the first chart sheet, and removes its chart control using ShapeCollection.RemoveAt(0) with Aspose.Cells. | Generate a snippet that iterates through worksheets, identifies a chart sheet, and deletes a specific shape by index from its Shapes collection using Aspose.Cells. | Show how to adapt the example to remove a chart shape by name instead of index in a chart sheet with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# remove chart shape from chart sheet | how to delete a chart control on a chart sheet using Aspose.Cells | ShapeCollection.RemoveAt example for Excel chart sheet in C# | programmatically remove chart sheet shapes with Aspose.Cells | C# Aspose.Cells delete first shape on a chart sheet
// Tags: chart shape removal Aspose.Cells C# | ShapeCollection.RemoveAt chart sheet | delete chart control Aspose.Cells | Aspose.Cells chart sheet shape management | C# remove chart shape from Excel workbook

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel file, finds the first chart sheet, accesses its Shapes collection, removes the first chart control using ShapeCollection.RemoveAt(0), and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the chart sheet
            Workbook workbook = new Workbook(inputPath);

            // Find the first chart sheet in the workbook
            int chartSheetIndex = -1;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].Type == SheetType.Chart)
                {
                    chartSheetIndex = i;
                    break;
                }
            }

            if (chartSheetIndex >= 0)
            {
                // Get the chart sheet worksheet
                Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];

                // Access the shape collection of the chart sheet
                ShapeCollection shapes = chartSheet.Shapes;

                // Remove the first shape (chart control) if any exist
                if (shapes.Count > 0)
                {
                    shapes.RemoveAt(0); // Adjust the index as needed
                }
            }
            else
            {
                Console.WriteLine("No chart sheet found in the workbook.");
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
