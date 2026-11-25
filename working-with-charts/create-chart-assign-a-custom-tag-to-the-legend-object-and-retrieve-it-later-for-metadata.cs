// Title: Create a column chart with Aspose.Cells for .NET, add a custom tag to its legend via workbook custom properties, and read the tag after saving
// Common Searches: Aspose.Cells C# add custom metadata to chart legend and read it later | store and retrieve custom document property for Excel chart legend using Aspose.Cells .NET | persist legend tag in Excel file created with Aspose.Cells column chart | how to tag a chart legend with a custom string in Aspose.Cells for .NET | read custom legend tag after reloading workbook Aspose.Cells C#
// Tags: Aspose.Cells chart legend custom property | column chart legend metadata storage | Excel workbook legend property persistence | retrieve legend tag after workbook reload | C# Aspose.Cells chart legend tagging

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartLegendTagDemo
{
    // The example creates a new Workbook, adds a column chart with sample data, accesses the chart's Legend object, stores a custom string in the workbook's CustomDocumentProperties under the name 'LegendTag', saves the file, and then reads the property back both before and after reloading the workbook to demonstrate that the legend tag is persisted.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "ChartWithLegendTag.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // 3. Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // 4. Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // 5. Access the legend object
                Legend legend = chart.Legend;

                // 6. Store custom metadata in workbook's custom document properties
                string customTag = "ReportSection:SalesSummary;CreatedBy:ChartLegendTagDemo";
                workbook.CustomDocumentProperties.Add("LegendTag", customTag);

                // Optional: customize the appearance of the legend
                legend.Position = LegendPositionType.Bottom;
                legend.Font.Size = 12;
                legend.Font.IsBold = true;

                // 7. Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");

                // 8. Retrieve the custom tag from the workbook's custom document properties
                string retrievedTag = workbook.CustomDocumentProperties["LegendTag"]?.Value?.ToString();
                if (!string.IsNullOrEmpty(retrievedTag))
                {
                    Console.WriteLine("Retrieved Legend Tag: " + retrievedTag);
                }
                else
                {
                    Console.WriteLine("No tag was found on the legend.");
                }

                // 9. Demonstrate that the tag is persisted after reloading the workbook.
                if (File.Exists(outputPath))
                {
                    Workbook reloaded = new Workbook(outputPath);
                    string reloadedTag = reloaded.CustomDocumentProperties["LegendTag"]?.Value?.ToString();
                    Console.WriteLine("Tag after reloading workbook: " + (reloadedTag ?? "null"));
                }
                else
                {
                    Console.WriteLine($"File \"{outputPath}\" not found for reloading.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
