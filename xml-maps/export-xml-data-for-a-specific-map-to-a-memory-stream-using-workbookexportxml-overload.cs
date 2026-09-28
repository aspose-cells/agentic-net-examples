// Title: Export the first map chart from an Excel workbook to an XML file using the Workbook.ExportXml overload in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, finds the first map chart on the first worksheet, and uses Workbook.ExportXml to write the chart's XML to a temporary file, then reads the file into a string and removes it. | Show how to add error handling for missing workbook files, absent worksheets, or no map charts when exporting map XML with Aspose.Cells.
// Common Searches: Aspose.Cells export map chart XML to file example | C# Workbook.ExportXml overload for map charts | How to get XML of a specific map chart from Excel using Aspose.Cells | Export chart of type Map to XML with Aspose.Cells .NET
// Tags: export map chart to XML Aspose.Cells | Workbook.ExportXml overload usage | temporary XML file handling C# | read exported XML content Aspose.Cells | locate first map chart worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Loads a workbook, searches the first worksheet for a map chart, exports that chart's XML to a temporary file via Workbook.ExportXml, reads the XML into a string, outputs it, and finally deletes the temporary file.
class ExportMapToXml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains at least one map chart.
            using (Workbook workbook = new Workbook(inputPath))
            {
                // Ensure the workbook has at least one worksheet.
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Error: The workbook does not contain any worksheets.");
                    return;
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Find the first map chart in the worksheet.
                Chart mapChart = null;
                foreach (Chart chart in sheet.Charts)
                {
                    if (chart.Type == ChartType.Map)
                    {
                        mapChart = chart;
                        break;
                    }
                }

                if (mapChart == null)
                {
                    Console.WriteLine("Error: No map charts found in the first worksheet.");
                    return;
                }

                // Use the map chart's name for export.
                string mapName = mapChart.Name;

                // Export the map to a temporary XML file.
                string tempXmlPath = Path.Combine(Path.GetTempPath(), $"ExportedMap_{Guid.NewGuid()}.xml");
                workbook.ExportXml(tempXmlPath, mapName);

                // Read and display the exported XML content.
                string xmlContent = File.ReadAllText(tempXmlPath);
                Console.WriteLine(xmlContent);

                // Clean up the temporary file.
                if (File.Exists(tempXmlPath))
                {
                    File.Delete(tempXmlPath);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
