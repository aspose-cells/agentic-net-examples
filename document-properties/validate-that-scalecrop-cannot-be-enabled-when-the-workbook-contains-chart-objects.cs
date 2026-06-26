using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Properties;

namespace AsposeCellsScaleCropValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Validate ScaleCrop setting: it must not be enabled when any chart exists
            bool workbookHasCharts = false;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Charts.Count > 0)
                {
                    workbookHasCharts = true;
                    break;
                }
            }

            BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

            if (workbookHasCharts)
            {
                // ScaleCrop cannot be enabled; keep it false and inform the user
                properties.ScaleCrop = false;
                Console.WriteLine("ScaleCrop was not enabled because the workbook contains chart objects.");
            }
            else
            {
                // No charts, safe to enable ScaleCrop
                properties.ScaleCrop = true;
                Console.WriteLine("ScaleCrop enabled: " + properties.ScaleCrop);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ScaleCropValidationResult.xlsx", SaveFormat.Xlsx);
        }
    }
}