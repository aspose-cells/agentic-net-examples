// Title: Clone a Chart, Adjust Its Secondary Axis, and Insert the Copy into a New Worksheet – Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to create a workbook with sample data, add a column chart that uses a secondary value axis, clone the worksheet containing the chart, retrieve the cloned chart, customize its secondary axis title, minimum, maximum and major unit, optionally reposition the chart on the new sheet, and save the result as an Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart clone C# | secondary axis Aspose.Cells | copy worksheet with chart .NET | duplicate Excel chart C# | move chart Aspose.Cells | Chart.Move method | AddCopy worksheet Aspose.Cells | C# Excel chart secondary value axis
// Common Searches: Aspose.Cells clone chart C# | How to copy a worksheet that contains a chart in Aspose.Cells | Change secondary axis of a cloned chart Aspose.Cells | Move chart to another sheet using Aspose.Cells for .NET | Set secondary axis range for Excel chart with Aspose.Cells
// Developer Intent: The developer wants to duplicate an existing chart, modify its secondary axis settings, and place the duplicated chart on a separate worksheet using Aspose.Cells for .NET.
// Use Cases: Create a report that shows the original chart and a version with a different secondary axis scale on separate sheets for side‑by‑side comparison. | Automate the replication of charts across multiple worksheets while customizing each copy’s secondary axis title and range to match distinct data sets. | Build a template workbook that copies a chart to a new sheet, repositions it, and applies a tailored secondary axis configuration for downstream processing.
// AI Prompts: Generate C# code with Aspose.Cells that clones a chart from one worksheet, changes the secondary axis title, min, max, and major unit, and inserts the clone into a new sheet. | Show how to copy a worksheet containing a chart, retrieve the duplicated chart, adjust its secondary axis parameters, and move the chart to a different location on the target sheet using Aspose.Cells for .NET. | Explain step‑by‑step how to duplicate a chart, customize its secondary value axis range, and save the workbook with the modified chart in Aspose.Cells (C#).

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example demonstrates how to create a workbook with sample data, add a column chart that uses a secondary value axis, clone the worksheet containing the chart, retrieve the cloned chart, customize its secondary axis title, minimum, maximum and major unit, optionally reposition the chart on the new sheet, and save the result as an Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate sample data for the chart
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["A2"].PutValue("A");
        sourceSheet.Cells["A3"].PutValue("B");
        sourceSheet.Cells["A4"].PutValue("C");

        sourceSheet.Cells["B1"].PutValue("Series 1");
        sourceSheet.Cells["B2"].PutValue(100);
        sourceSheet.Cells["B3"].PutValue(200);
        sourceSheet.Cells["B4"].PutValue(300);

        sourceSheet.Cells["C1"].PutValue("Series 2");
        sourceSheet.Cells["C2"].PutValue(5000);
        sourceSheet.Cells["C3"].PutValue(3000);
        sourceSheet.Cells["C4"].PutValue(1000);

        // Add the original chart
        int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart originalChart = sourceSheet.Charts[chartIndex];
        originalChart.NSeries.Add("B2:B4", true);
        originalChart.NSeries.Add("C2:C4", true);
        originalChart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        originalChart.NSeries[1].PlotOnSecondAxis = true;

        // Configure the secondary axis of the original chart (optional)
        Axis originalSecAxis = originalChart.SecondValueAxis;
        originalSecAxis.Title.Text = "Original Secondary Axis";
        originalSecAxis.MinValue = 0;
        originalSecAxis.MaxValue = 6000;
        originalSecAxis.MajorUnit = 1000;

        // Clone the chart by copying the entire worksheet that contains it
        int copiedSheetIndex = workbook.Worksheets.AddCopy("Source"); // uses AddCopy(string) rule
        Worksheet clonedSheet = workbook.Worksheets[copiedSheetIndex];
        clonedSheet.Name = "ClonedChartSheet";

        // Retrieve the cloned chart (same index as in the source sheet)
        Chart clonedChart = clonedSheet.Charts[chartIndex];

        // Modify the secondary axis settings of the cloned chart
        Axis clonedSecAxis = clonedChart.SecondValueAxis;
        clonedSecAxis.Title.Text = "Cloned Secondary Axis";
        clonedSecAxis.MinValue = 0;
        clonedSecAxis.MaxValue = 8000;
        clonedSecAxis.MajorUnit = 2000;

        // Optionally reposition the cloned chart on the new sheet
        clonedChart.Move(5, 10, 20, 18); // uses Chart.Move method

        // Save the workbook (uses the provided save rule)
        workbook.Save("ClonedChartDemo.xlsx");
    }
}
