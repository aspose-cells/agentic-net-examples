// Title: Aspose.Cells for .NET – Add a Stacked Bar Chart as a Gantt Chart to an Existing XLSX Workbook
// Description: Load an XLSX file, insert a stacked bar chart, define start and duration series, hide the start series using GapWidth and Overlap, set task names as categories, and save the workbook with a ready‑to‑use Gantt chart.
// Keywords: Aspose.Cells Gantt chart .NET | stacked bar chart Aspose.Cells | hide series Aspose chart | chart GapWidth Overlap Aspose | create Gantt chart programmatically | C# Excel chart automation
// Common Searches: how to create a Gantt chart with Aspose.Cells C# | Aspose.Cells stacked bar chart hide start series | set overlap and gap width for Gantt chart Aspose | add chart to existing workbook Aspose.Cells | C# generate Gantt chart from task data
// Developer Intent: Programmatically generate a Gantt‑style chart by adding and formatting a stacked bar chart in an existing Excel workbook using Aspose.Cells for .NET.
// Use Cases: Transform a task‑list worksheet into a visual Gantt chart for project status reports. | Automate Gantt chart creation across multiple project files without manual chart design. | Embed a printable Gantt chart directly into a financial or resource‑allocation workbook.
// AI Prompts: Write C# code with Aspose.Cells that loads an XLSX file, adds a stacked bar chart, configures start and duration series, and uses GapWidth/Overlap to produce a Gantt chart. | Explain how to hide the start series in an Aspose.Cells stacked bar chart while preserving its data for positioning. | Show how to make the task, start date, and duration ranges dynamic when building a Gantt chart with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLSX file, insert a stacked bar chart, define start and duration series, hide the start series using GapWidth and Overlap, set task names as categories, and save the workbook with a ready‑to‑use Gantt chart.
class GanttChartExample
{
    public static void Main()
    {
        // Load an existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (where data is stored)
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Assume the worksheet already contains the data:
        //   Column A : Task names (categories)
        //   Column B : Start values (e.g., start day)
        //   Column C : Duration values (e.g., length of task)
        // -------------------------------------------------

        // Add a stacked bar chart to the sheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Explicitly set the chart type (optional, already BarStacked)
        chart.Type = ChartType.BarStacked;

        // Add the two series required for a Gantt chart
        // First series – Start (will be hidden later)
        chart.NSeries.Add("B2:B6", true);
        // Second series – Duration (visible part of the bar)
        chart.NSeries.Add("C2:C6", true);

        // Set the category (task names) for the X‑axis
        chart.NSeries.CategoryData = "A2:A6";

        // -------------------- Gantt formatting --------------------
        // Reduce the gap between bars for a tighter look
        chart.GapWidth = 150;               // 150% of default gap

        // Stack the two series without any space between them
        chart.NSeries[0].Overlap = -100;    // full overlap

        // Hide the start series by making its gap width zero
        // (the series still exists for positioning but is not drawn)
        chart.NSeries[0].GapWidth = 0;

        // ---------------------------------------------------------

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
