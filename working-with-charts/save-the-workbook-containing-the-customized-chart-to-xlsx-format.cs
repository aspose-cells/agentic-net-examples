// Title: Saving an Aspose.Cells workbook that contains a customized column chart to XLSX using C#
// Common Searches: C# code to export an Aspose.Cells workbook with a column chart to an .xlsx file | How to use Aspose.Cells SaveFormat.Xlsx to write a chart workbook in .NET | Aspose.Cells example for saving a workbook with customized chart colors to XLSX | Write a workbook containing a column chart to disk with Aspose.Cells in C#
// Tags: Aspose.Cells save workbook with chart to XLSX | C# column chart customization Aspose.Cells | Aspose.Cells SaveFormat.Xlsx usage | export chart workbook Aspose.Cells .NET | custom series color Aspose.Cells chart

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, populates cells A1:B4 with sample data, adds a column chart, customizes its title and series color, and then saves the workbook as 'CustomizedChart.xlsx' in XLSX format using Aspose.Cells' Workbook.Save method with SaveFormat.Xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet (position: rows 5-20, columns 0-10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Sample Column Chart";

            // Add series data (values from B2:B4) and category labels (A2:A4)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Optional: customize series appearance
            chart.NSeries[0].Area.ForegroundColor = Color.Blue;

            // Determine output file path
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "CustomizedChart.xlsx");

            // Save the workbook in XLSX format
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
