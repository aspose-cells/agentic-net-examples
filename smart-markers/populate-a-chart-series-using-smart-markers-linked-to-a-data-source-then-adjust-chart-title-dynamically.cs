// Title: Create an Excel column chart with smart markers bound to a List<T> and set the chart title at runtime using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells WorkbookDesigner to populate a column chart series from a List<Record> via smart markers and then assign a scalar value to the chart title after processing. | Show how to define smart‑marker placeholders for a data range, bind a collection as a data source, generate the chart, and update the chart title dynamically with Aspose.Cells.
// Common Searches: Aspose.Cells C# bind List<T> to smart markers for chart series | How to set Excel chart title dynamically after WorkbookDesigner.Process in .NET | Smart markers example for column chart with data source list in Aspose.Cells | Create chart with smart markers and dynamic title using Aspose.Cells for .NET
// Tags: Aspose.Cells WorkbookDesigner smart markers list binding | C# column chart series from smart markers | dynamic chart title Aspose.Cells | populate Excel chart using List<T> smart markers | smart markers data source Excel .xlsx generation

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SmartMarkersChartDemo
{
    // Simple data class for smart markers
    // The sample creates a workbook, adds headers, inserts smart‑marker placeholders for Category and Value, binds a List<Record> as the data source, sets a scalar ReportTitle, processes all smart markers with WorkbookDesigner, creates a column chart that references the smart‑marker range, assigns the processed ReportTitle to the chart title, and saves the file as SmartMarkersChart.xlsx.
    public class Record
    {
        public string Category { get; set; }
        public int Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Insert smart markers for data rows (5 rows)
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue("&=$Category");
                sheet.Cells[$"B{i}"].PutValue("&=$Value");
            }

            // Add a column chart that references the smart‑marker range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);          // Values
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";     // Categories

            // Prepare data source for smart markers
            List<Record> data = new List<Record>
            {
                new Record { Category = "Jan", Value = 120 },
                new Record { Category = "Feb", Value = 150 },
                new Record { Category = "Mar", Value = 180 },
                new Record { Category = "Apr", Value = 200 },
                new Record { Category = "May", Value = 170 }
            };

            // Title that will be set dynamically after processing
            string reportTitle = "Monthly Sales Report";

            // Set up WorkbookDesigner with data sources
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource("Data", data);          // Smart markers will use this list
            designer.SetDataSource("ReportTitle", reportTitle); // Scalar for title

            // Process all smart markers in the workbook
            designer.Process();

            // Adjust chart title dynamically using the scalar value
            chart.Title.Text = reportTitle;

            // Save the workbook
            workbook.Save("SmartMarkersChart.xlsx");
        }
    }
}
