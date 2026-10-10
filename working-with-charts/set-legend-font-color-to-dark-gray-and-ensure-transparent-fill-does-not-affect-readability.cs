// Title: How to set a chart legend's font color to dark gray and make its background transparent using Aspose.Cells for .NET
// AI Prompts: Update the first chart in an Excel workbook so its legend text uses DarkGray and the legend area has no fill using Aspose.Cells C#. | Apply a transparent fill to a chart legend while changing the legend text color to dark gray with Aspose.Cells for .NET. | Programmatically change a chart legend's text color to dark gray and remove its background fill, then save the workbook using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# set chart legend text color to dark gray | make chart legend background transparent Aspose.Cells .NET | remove legend fill type in Excel chart using Aspose.Cells | change legend text color and hide legend background in C# Excel library | how to format chart legend appearance with Aspose.Cells
// Tags: chart legend font color Aspose.Cells | transparent legend fill Aspose.Cells | set legend fill type none .NET | Excel chart legend styling C# | Aspose.Cells chart legend formatting

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, checks for a chart on the first worksheet, sets the legend's text color to DarkGray, removes the legend's background fill by setting FillType to None, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists; create a placeholder workbook if it does not.
        if (!File.Exists(inputPath))
        {
            try
            {
                var placeholder = new Workbook();
                placeholder.Save(inputPath);
                Console.WriteLine($"Created placeholder workbook at '{inputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create placeholder workbook: {ex.Message}");
                return;
            }
        }

        Workbook workbook;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook '{inputPath}': {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Modify the first chart's legend if a chart exists
        if (sheet.Charts.Count > 0)
        {
            try
            {
                Chart chart = sheet.Charts[0];

                // Set legend font color to dark gray
                chart.Legend.Font.Color = Color.DarkGray;

                // Make legend background transparent by setting no fill
                chart.Legend.Area.FillFormat.FillType = FillType.None;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error modifying chart: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("No charts found on the first worksheet.");
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
