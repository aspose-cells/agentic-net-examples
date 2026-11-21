// Title: Make chart legend entries fully transparent in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to set Excel chart legend entries transparent with Aspose.Cells | Aspose.Cells .NET how to hide legend background color in a chart | remove legend entry fill color in Excel chart using Aspose.Cells C# | make chart legend font invisible Aspose.Cells example | transparent legend entries Aspose.Cells chart manipulation C#
// Tags: transparent legend entry font Aspose.Cells | chart legend background removal C# Aspose.Cells | set legend entry color to transparent Aspose.Cells .NET | modify Excel chart legend fill Aspose.Cells | Aspose.Cells legend transparency example | Excel chart legend styling Aspose.Cells C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing workbook, accesses the first worksheet's first chart, iterates through each legend entry, and sets its font color to fully transparent using Color.FromArgb(0,0,0,0) before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found on the first worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Iterate through all legend entries and make them fully transparent
            foreach (LegendEntry legendEntry in chart.Legend.LegendEntries)
            {
                // Set the legend entry font color to fully transparent
                legendEntry.Font.Color = Color.FromArgb(0, 0, 0, 0);
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
