using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkersChartDemo
{
    // Simple data class for the smart markers
    public class Person
    {
        public string Name { get; set; } = null!;   // initialized by data source
        public double Sales { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Prepare sample data source
                List<Person> persons = new List<Person>
                {
                    new Person { Name = "Apple",  Sales = 1200 },
                    new Person { Name = "Banana", Sales = 850 },
                    new Person { Name = "Cherry", Sales = 430 },
                    new Person { Name = "Date",   Sales = 670 }
                };

                // 3. Insert smart markers that will be replaced by the data source
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("&=$Name");
                sheet.Cells["B2"].PutValue("&=$Sales");

                // Define the range that contains the smart markers
                AsposeRange smRange = sheet.Cells.CreateRange("A2:B2");
                smRange.Name = "_CellsSmartMarkers";

                // 4. Create a chart that will use the cells filled by smart markers
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // The data range will be populated after the smart markers are processed
                chart.NSeries.Add("=Sheet1!$B$2:$B$5", true);
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

                // Enable data labels for better visibility
                chart.NSeries[0].DataLabels.ShowValue = true;

                // 5. Use WorkbookDesigner to bind the data source and process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Data", persons);
                designer.Process(); // processes all smart markers in the workbook

                // 6. Dynamically adjust the chart title after data has been populated
                chart.Title.Text = $"Sales Report – {persons.Count} Products";
                chart.Title.OverLay = true; // overlay title without resizing the chart

                // 7. Save the workbook (ensure the directory exists)
                string outputPath = "SmartMarkersChart.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}