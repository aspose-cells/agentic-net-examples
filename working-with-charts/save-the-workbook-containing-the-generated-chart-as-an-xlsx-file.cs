// Title: Create a column chart in a workbook and save it as an XLSX file with Aspose.Cells for .NET (C#)
// Common Searches: how to generate a column chart and export the workbook to XLSX using Aspose.Cells in C# | Aspose.Cells C# save workbook with embedded chart to .xlsx | C# code to add a chart to a worksheet and save as XLSX with Aspose.Cells | Aspose.Cells example for creating a chart and saving the file in XLSX format | save workbook containing chart to xlsx using Aspose.Cells .NET API
// Tags: column chart creation Aspose.Cells C# | workbook save to xlsx Aspose.Cells | chart data binding Aspose.Cells | Aspose.Cells chart export example | C# Aspose.Cells chart workbook generation

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSave
{
    // The sample program creates a new Workbook, fills cells A1:B4 with sample data, adds a column chart linked to that data, sets a chart title, and then saves the workbook as an XLSX file named "ChartWorkbook.xlsx" using the Aspose.Cells for .NET API.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", false);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: give the chart a title
                chart.Title.Text = "Sample Column Chart";

                // Save the workbook as XLSX
                string outputPath = "ChartWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook with chart saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating or saving the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
