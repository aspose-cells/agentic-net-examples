using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetDataLabelBorderThickness
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked bar chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a stacked bar chart (BarStacked is the correct enum value)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels for the first series
            DataLabels labels = chart.NSeries[0].DataLabels;
            labels.ShowValue = true;               // Ensure labels are displayed
            labels.Border.IsVisible = true;        // Make the border visible
            labels.Border.WeightPt = 2.0;           // Set border thickness to 2 points

            // Save the workbook
            string outputPath = "StackedBarDataLabelBorder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}