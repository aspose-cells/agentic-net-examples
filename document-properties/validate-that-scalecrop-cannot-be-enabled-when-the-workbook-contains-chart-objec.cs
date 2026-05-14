using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsValidation
{
    public class ScaleCropChartValidation
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for a chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Use a custom document property to represent ScaleCrop
            var customProps = workbook.CustomDocumentProperties;
            const string scaleCropName = "ScaleCrop";

            // Initialize ScaleCrop to true (attempt to enable)
            if (!customProps.Contains(scaleCropName))
                customProps.Add(scaleCropName, true);
            else
                customProps[scaleCropName].Value = true;

            // Validation: ScaleCrop must be false if any chart objects are present
            bool hasChart = false;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Charts.Count > 0)
                {
                    hasChart = true;
                    break;
                }
            }

            bool scaleCrop = (bool)customProps[scaleCropName].Value;

            if (hasChart && scaleCrop)
            {
                // Reset the property and inform the user
                customProps[scaleCropName].Value = false;
                Console.WriteLine("ScaleCrop cannot be enabled because the workbook contains chart objects. It has been reset to false.");
            }
            else
            {
                Console.WriteLine("ScaleCrop setting is valid.");
            }

            // Save the workbook
            workbook.Save("ScaleCropValidationResult.xlsx", SaveFormat.Xlsx);
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}