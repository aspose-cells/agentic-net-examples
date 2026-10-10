// Title: How to make Aspose.Cells chart data labels inherit number formats from source worksheet cells in C#
// AI Prompts: Write C# code using Aspose.Cells that creates a column chart and configures its data labels to automatically use the number format of the cells they reference. | Update the given Aspose.Cells example so that data label formatting is inherited for percentage and currency values displayed in the chart. | Describe why the current Aspose.Cells API cannot bind data label number formats to source cells and propose a manual workaround.
// Common Searches: Aspose.Cells chart data label format same as cell format C# | inherit number format from worksheet cells for chart labels Aspose.Cells | dynamic formatting of chart data labels based on source cell in .NET | link chart data label to cell number format Aspose.Cells example
// Tags: Aspose.Cells chart data label formatting | C# link chart label to cell number format | dynamic number format inheritance Aspose.Cells | column chart data labels Aspose.Cells | worksheet cell formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads or creates a workbook, applies percentage and currency number formats to cells A1 and B1, adds a column chart that references those cells, enables data label values, and saves the file; the Aspose.Cells version used does not provide a property to bind data label number formats to the source cells, so formatting must be applied manually.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate A1 with a percentage value and apply percentage format.
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue(0.2567);
            Style styleA1 = cellA1.GetStyle();
            styleA1.Number = 10; // Percentage format.
            cellA1.SetStyle(styleA1);

            // Populate B1 with a currency value and apply currency format.
            Cell cellB1 = sheet.Cells["B1"];
            cellB1.PutValue(15890);
            Style styleB1 = cellB1.GetStyle();
            styleB1.Number = 3; // Currency format.
            cellB1.SetStyle(styleB1);

            // Add a column chart to the worksheet.
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add a series that references the formatted cells.
            chart.NSeries.Add("A1:B1", true);

            // Enable data labels and show values.
            chart.NSeries[0].DataLabels.ShowValue = true;
            // Linking number format to source is not available in this version; omitted.

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
