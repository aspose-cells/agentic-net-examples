// Title: Copy a chart, modify its title, and place it on a summary worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy chart from one worksheet to another and rename the title | programmatically duplicate a column chart on a summary sheet with Aspose.Cells | Aspose.Cells .NET create a chart on a data sheet then replicate it on another worksheet | change chart title after copying a chart shape using Aspose.Cells C# example | preserve chart style while copying chart between worksheets in Aspose.Cells
// Tags: copy chart between worksheets Aspose.Cells .NET | update chart title C# Aspose.Cells | duplicate chart shape summary worksheet | recreate chart with same data series Aspose.Cells | preserve chart style Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data to a 'Data' worksheet, inserts a column chart, then adds a 'Summary' worksheet and recreates the same chart there using the original data ranges and style, updates the chart title to indicate it is a copied version, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and add a source worksheet with data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the chart
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["B4"].PutValue(30);

            // ------------------------------------------------------------
            // 2. Add a chart to the source worksheet
            // ------------------------------------------------------------
            int sourceChartIdx = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart sourceChart = dataSheet.Charts[sourceChartIdx];
            sourceChart.NSeries.Add("B2:B4", true);               // Values
            sourceChart.NSeries.CategoryData = "A2:A4";           // Categories
            sourceChart.Title.Text = "Original Chart";

            // ------------------------------------------------------------
            // 3. Create a summary worksheet where the chart will be copied
            // ------------------------------------------------------------
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");

            // ------------------------------------------------------------
            // 4. Recreate the chart on the summary sheet with the same data and style
            // ------------------------------------------------------------
            int summaryChartIdx = summarySheet.Charts.Add(sourceChart.Type, 2, 0, 15, 8);
            Chart summaryChart = summarySheet.Charts[summaryChartIdx];

            // Copy the data series (values range) – using the same range strings as the source
            summaryChart.NSeries.Add("B2:B4", true);
            // Copy the category data for the whole series collection
            summaryChart.NSeries.CategoryData = "A2:A4";

            // Copy visual style (if any)
            summaryChart.Style = sourceChart.Style;

            // Change the title to indicate this is the copied version
            summaryChart.Title.Text = "Copied Chart – Summary";

            // ------------------------------------------------------------
            // 5. Save the workbook
            // ------------------------------------------------------------
            string outFile = "ChartCopySummary.xlsx";
            workbook.Save(outFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully as \"{outFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
