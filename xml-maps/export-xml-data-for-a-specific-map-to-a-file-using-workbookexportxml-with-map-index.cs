// Title: Export a specific map chart’s XML data to a file with Workbook.ExportXml in Aspose.Cells for .NET
// Common Searches: how to export a map chart to XML using Aspose.Cells C# | Aspose.Cells example for exporting a specific XML map by chart name | C# code to save Excel map data as an XML file with Aspose.Cells | retrieve the first map chart in a worksheet and export its XML | Workbook.ExportXml method usage with map chart name in .NET
// Tags: Workbook.ExportXml chart name parameter | extract map chart XML data | identify map chart index in worksheet | Aspose.Cells export XML map | C# write Excel map to XML file

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts; // For Chart and ChartType

// The sample loads an Excel workbook, locates the first map chart on the first worksheet, and uses Workbook.ExportXml with the chart's name to write the map's XML representation to a specified file, including error handling for missing files or absent map charts.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "mapData.xml";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains a map chart
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Find the first map chart in the worksheet
            Chart mapChart = sheet.Charts.FirstOrDefault(c => c.Type == ChartType.Map);
            if (mapChart == null)
            {
                Console.WriteLine("Error: No map charts found in the first worksheet.");
                return;
            }

            // Export the map data to an XML file using the chart's name
            workbook.ExportXml(outputPath, mapChart.Name);
            Console.WriteLine($"Map data exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
