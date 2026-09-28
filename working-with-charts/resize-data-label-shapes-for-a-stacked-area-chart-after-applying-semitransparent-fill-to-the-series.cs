// Title: Resize data label font size in a stacked area chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a stacked area chart, enables data labels for the first series, and sets the label font size to 14 points. | Show how to apply a semi‑transparent fill to a stacked area chart series in Aspose.Cells (when AreaFormat is available) and then enlarge the data label font. | Provide an example that saves the workbook after customizing chart data label appearance with Aspose.Cells in a .NET application. | Demonstrate adjusting only the first series' data label font size without affecting other series in an Aspose.Cells chart.
// Common Searches: Aspose.Cells C# increase data label font size for stacked area chart | How to set semi transparent fill on chart series using Aspose.Cells .NET | Resize chart data label appearance in Excel workbook with Aspose.Cells | Enable and format data labels for first series in Aspose.Cells stacked area chart | C# Aspose.Cells chart customization data label shape size
// Tags: Aspose.Cells stacked area chart data label font size | Aspose.Cells apply semi transparent fill to series | C# resize chart data label appearance | Aspose.Cells chart label formatting example | Excel workbook chart customization Aspose.Cells .NET | Aspose.Cells chart series visual style

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, adds a stacked area chart with sample data, enables data labels for the first series, sets the label font size to 14 points, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the stacked area chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(40);
            sheet.Cells["B4"].PutValue(50);
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(30);
            sheet.Cells["C4"].PutValue(40);

            // Add a stacked area chart
            int chartIndex = sheet.Charts.Add(ChartType.AreaStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:C4", true);

            // NOTE: Older Aspose.Cells versions may not expose AreaFormat on Series.
            // If available, you can set a semi‑transparent fill here.
            // The following loop is omitted to maintain compatibility.

            // Show data labels for the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Increase font size of data labels (makes them appear larger)
            chart.NSeries[0].DataLabels.Font.Size = 14;

            // Save the workbook
            string outputPath = "StackedAreaChart_WithResizedDataLabels.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
