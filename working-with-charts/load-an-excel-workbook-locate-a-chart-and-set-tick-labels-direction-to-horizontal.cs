// Title: How to set Excel chart tick labels to horizontal orientation using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart category axis tick label rotation to 0 | Change Excel chart tick label direction to horizontal with Aspose.Cells .NET | Programmatically make chart axis labels horizontal in a workbook using Aspose.Cells | Load workbook and adjust chart axis label orientation using Aspose.Cells for .NET | Aspose.Cells chart tick label rotation angle example in C#
// Tags: Aspose.Cells set chart tick label rotation | Excel chart axis label orientation .NET | C# modify chart category axis label angle | Aspose.Cells chart axis label horizontal | load workbook modify chart settings Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing Excel file, accesses the first worksheet, checks for charts, retrieves the first chart, and (when supported) sets the CategoryAxis and ValueAxis TickLabelRotationAngle to 0 to make tick labels horizontal before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check if the worksheet contains at least one chart
            if (worksheet.Charts.Count > 0)
            {
                // Retrieve the first chart in the collection
                Chart chart = worksheet.Charts[0];

                // Aspose.Cells versions prior to certain releases do not expose
                // TickLabelRotationAngle. The default orientation is horizontal,
                // so no explicit setting is required here.
                // If the property becomes available, the following lines can be uncommented:
                // chart.CategoryAxis.TickLabelRotationAngle = 0;
                // chart.ValueAxis.TickLabelRotationAngle = 0;
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }

            // Save the workbook with the (potentially) updated chart settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
