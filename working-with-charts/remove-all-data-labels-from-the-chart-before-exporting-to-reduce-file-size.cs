// Title: Remove all data labels from every chart series in an Aspose.Cells workbook using C# to shrink the exported XLSX file
// AI Prompts: Write C# code that iterates over each worksheet and each chart in an Aspose.Cells workbook, sets Series.DataLabels.IsDeleted = true for every series, and saves the workbook as XLSX. | Show how to disable chart data labels for all series with Aspose.Cells for .NET to minimize the size of the generated Excel file.
// Common Searches: asp.net remove chart data labels before saving workbook Aspose.Cells | c# delete all data labels from charts in Aspose.Cells to reduce file size | how to hide series data labels in Aspose.Cells chart programmatically | optimize Excel output size by removing chart labels using Aspose.Cells C#
// Tags: Aspose.Cells chart series label suppression C# | DataLabels.IsDeleted usage Aspose.Cells | reduce exported XLSX size Aspose.Cells | loop through worksheets and charts Aspose.Cells | disable chart data labels programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates looping through all worksheets and charts in a workbook, marking each series' DataLabels as deleted, and saving the workbook as XLSX, which reduces the resulting file size.
    class RemoveDataLabelsDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels initially (optional, just to demonstrate removal)
            chart.NSeries[0].DataLabels.ShowValue = true;

            // ------------------------------------------------------------
            // Remove all data labels from every series in every chart
            // ------------------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    foreach (Series ser in ch.NSeries)
                    {
                        // Mark the DataLabels object as deleted – this removes all labels
                        ser.DataLabels.IsDeleted = true;
                    }
                }
            }

            // Save the workbook (export)
            workbook.Save("ChartWithoutDataLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
