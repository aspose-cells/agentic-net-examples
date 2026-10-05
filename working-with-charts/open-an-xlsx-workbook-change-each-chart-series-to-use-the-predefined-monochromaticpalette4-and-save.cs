// Title: Set every chart in an XLSX workbook to MonochromePalette4 using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing XLSX file with Aspose.Cells, loop through all worksheets and charts, assign ChartPaletteType.MonochromePalette4 to each chart's Palette property, and save the workbook. | Use C# reflection to set the Palette property on a chart when the property is not directly accessible in the current Aspose.Cells version. | Implement per‑chart error handling while applying a monochrome color scheme to all charts in a workbook.
// Common Searches: C# Aspose.Cells set chart palette to MonochromePalette4 for all charts in a workbook | apply monochrome color scheme to Excel charts using Aspose.Cells .NET | how to use reflection to set chart Palette property in Aspose.Cells when property is unavailable | iterate through worksheets and charts in Aspose.Cells to change chart colors
// Tags: Aspose.Cells set chart palette C# | MonochromePalette4 Excel chart styling | iterate worksheets charts Aspose.Cells | reflection assign chart property .NET | apply monochrome color scheme XLSX

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code loads 'input.xlsx', iterates over every worksheet and chart, uses reflection to assign ChartPaletteType.MonochromePalette4 to each chart's Palette property when possible, logs any per‑chart failures, and saves the modified workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook from the existing XLSX file
            Workbook workbook = new Workbook(inputFile);

            // Iterate through all worksheets and try to apply a monochrome palette to each chart
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Use reflection to set the Palette property if it exists (covers different library versions)
                        var chartType = chart.GetType();
                        var paletteProp = chartType.GetProperty("Palette");
                        if (paletteProp != null && paletteProp.CanWrite)
                        {
                            // Obtain the ChartPaletteType enum type via its full name
                            var enumType = Type.GetType("Aspose.Cells.Charts.ChartPaletteType");
                            if (enumType != null)
                            {
                                // Parse the desired enum value
                                var enumValue = Enum.Parse(enumType, "MonochromePalette4");
                                paletteProp.SetValue(chart, enumValue);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log but continue processing other charts
                        Console.WriteLine($"Failed to set palette for a chart: {ex.Message}");
                    }
                }
            }

            // Save the modified workbook to a new XLSX file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
