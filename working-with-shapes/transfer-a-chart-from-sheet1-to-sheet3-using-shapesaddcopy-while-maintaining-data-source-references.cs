using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class TransferChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Source sheet (Sheet1) --------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Populate sample data for the chart
            sheet1.Cells["A1"].PutValue("Category");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["A4"].PutValue("Cherry");

            sheet1.Cells["B1"].PutValue("Value");
            sheet1.Cells["B2"].PutValue(30);
            sheet1.Cells["B3"].PutValue(45);
            sheet1.Cells["B4"].PutValue(25);

            // Add a chart to Sheet1
            int chartIndex = sheet1.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart sourceChart = sheet1.Charts[chartIndex];
            sourceChart.NSeries.Add("B2:B4", true);
            // CategoryData may not be supported in all versions; omitted for compatibility

            // -------------------- Destination sheet (Sheet3) --------------------
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Get the chart shape (ChartShape) from the source chart
            Shape sourceChartShape = sourceChart.ChartObject;

            // Copy the chart shape to Sheet3 using Shapes.AddCopy
            Shape copiedChartShape = sheet3.Shapes.AddCopy(
                sourceChartShape,
                sourceChartShape.UpperLeftRow,
                sourceChartShape.UpperLeftColumn,
                sourceChartShape.LowerRightRow,
                sourceChartShape.LowerRightColumn);

            // The copied chart automatically appears in Sheet3.Charts collection
            Chart copiedChart = sheet3.Charts[0];

            // Verify that the data source still refers to Sheet1 (default behavior)
            Console.WriteLine("Original chart values formula: " + sourceChart.NSeries[0].Values);
            Console.WriteLine("Copied chart values formula: " + copiedChart.NSeries[0].Values);

            // Save the workbook
            string outputPath = "ChartTransferResult.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}