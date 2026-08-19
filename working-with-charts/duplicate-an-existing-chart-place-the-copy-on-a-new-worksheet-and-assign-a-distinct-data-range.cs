// Title: Copy a Chart to a New Worksheet and Assign a New Data Range with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a source column chart, duplicate the entire worksheet (including the chart) to a second sheet, and then reassign the copied chart to a different cell range and title before saving the file.
// Keywords: Aspose.Cells copy chart | duplicate chart .NET | chart data range Aspose.Cells | clone worksheet with chart | Aspose.Cells chart example C# | set chart source data programmatically | Aspose.Cells workbook automation
// Common Searches: How to copy a chart to another sheet using Aspose.Cells | Change data source of a duplicated chart in C# | Aspose.Cells duplicate worksheet and modify chart | Copy chart without losing formatting Aspose.Cells | Set new range for copied chart Aspose.Cells .NET
// Developer Intent: Create a chart copy on a separate worksheet and point it to a distinct data range.
// Use Cases: Produce side‑by‑side visual comparisons of original vs. updated metrics. | Generate regional reports where a template chart is cloned for each locale with its own data set. | Automate monthly reporting templates that replicate a master chart and bind each copy to month‑specific values.
// AI Prompts: Show C# code to copy only a chart (not the whole sheet) while preserving its style with Aspose.Cells. | Give an example of assigning a dynamic data range to a duplicated chart based on the number of rows. | Explain how to rename a copied chart and update its source range without affecting the original chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDuplication
{
    // Demonstrates how to create a workbook, add a source column chart, duplicate the entire worksheet (including the chart) to a second sheet, and then reassign the copied chart to a different cell range and title before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";

            // Populate sample data for the original chart
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["A4"].PutValue("C");
            sourceSheet.Cells["B4"].PutValue(30);

            // Add a chart to the source sheet
            int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart sourceChart = sourceSheet.Charts[chartIdx];
            sourceChart.SetChartDataRange("A1:B4", true);
            sourceChart.Title.Text = "Original Chart";

            // Duplicate the worksheet (including the chart) using AddCopy
            int copiedIndex = workbook.Worksheets.AddCopy("SourceSheet");
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "CopiedSheet";

            // Access the copied chart (assumes only one chart on the sheet)
            Chart copiedChart = copiedSheet.Charts[0];

            // Assign a distinct data range to the copied chart
            // Populate new data for the distinct range
            copiedSheet.Cells["A5"].PutValue("Category");
            copiedSheet.Cells["B5"].PutValue("Value");
            copiedSheet.Cells["A6"].PutValue("X");
            copiedSheet.Cells["B6"].PutValue(40);
            copiedSheet.Cells["A7"].PutValue("Y");
            copiedSheet.Cells["B7"].PutValue(50);
            copiedSheet.Cells["A8"].PutValue("Z");
            copiedSheet.Cells["B8"].PutValue(60);

            // Set the new data range for the copied chart
            copiedChart.SetChartDataRange("A5:B8", true);
            copiedChart.Title.Text = "Copied Chart with New Data";

            // Save the workbook
            workbook.Save("ChartDuplicationResult.xlsx");
        }
    }
}
