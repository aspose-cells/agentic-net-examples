using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (no template file needed, but check if a template path is provided)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked column chart
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

            // Add a stacked column chart (use ColumnStacked enum – correct name in Aspose.Cells)
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (both series)
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and configure them
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;   // Show percentages on data labels
            dataLabels.ShowValue = false;       // Hide raw values (optional)

            // Adjust the font size of the data labels
            dataLabels.Font.Size = 12;          // Desired font size
            dataLabels.ApplyFont();             // Apply the font settings to all child labels

            // Save the workbook
            string outputPath = "StackedColumnChart_PercentageLabels.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}