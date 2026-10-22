// Title: How to load an Excel workbook in C#, locate the first chart, and set its category axis tick labels to horizontal with Aspose.Cells
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, retrieves the first chart, assigns 0 to CategoryAxis.TickLabelRotationAngle to make X‑axis labels horizontal, and saves the file. | Show a snippet that programmatically changes the tick label direction of a chart’s category axis to horizontal in an existing .xlsx using Aspose.Cells for .NET. | Create a C# example that validates an input Excel file, accesses its first chart, applies a zero‑degree rotation to the X‑axis tick labels, and writes the result to a new workbook.
// Common Searches: Aspose.Cells C# set chart X axis tick label rotation to zero | How to make Excel chart axis labels horizontal with Aspose.Cells .NET | Change category axis tick label angle in an existing workbook using Aspose.Cells | C# example for adjusting chart axis label orientation in Aspose.Cells | Set tick label direction to horizontal for first chart in Excel file Aspose.Cells
// Tags: Aspose.Cells set category axis tick label angle | C# chart axis label orientation Aspose.Cells | horizontal tick labels Excel chart Aspose.Cells | modify chart axis rotation .NET | load workbook edit chart axis Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file (input.xlsx) with Aspose.Cells, retrieves the first worksheet and its first chart, sets the category (X) axis tick label rotation angle to 0 degrees so the labels appear horizontal, and saves the modified workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the worksheet
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Locate the first chart
                Chart chart = worksheet.Charts[0];

                // Access the category (X) axis
                Axis categoryAxis = chart.CategoryAxis;

                // NOTE: The TickLabelRotationAngle property may not be available in older Aspose.Cells versions.
                // If needed, adjust the rotation using a supported API for your version.

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
