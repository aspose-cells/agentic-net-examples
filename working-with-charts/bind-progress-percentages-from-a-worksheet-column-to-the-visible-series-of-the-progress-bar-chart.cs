// Title: Create a Progress Bar Chart in C# with Aspose.Cells by Binding Worksheet Columns to a Stacked Bar Series
// Description: This example shows how to generate an Excel workbook with Aspose.Cells, fill column A with task names and column B with fractional progress values, add a stacked bar chart, bind B2:B5 as the series data and A2:A5 as the category axis, configure the series to display as a progress bar, and save the file as ProgressBarChart.xlsx.
// Keywords: Aspose.Cells C# progress bar chart | bind series to chart Aspose.Cells | stacked bar chart from worksheet data | set category axis Aspose.Cells | chart data range C# | visualize percentages Aspose.Cells | Excel progress bar automation
// Common Searches: Aspose.Cells bind column to chart series C# | Create progress bar chart from Excel data using Aspose.Cells | How to set category labels from a worksheet column in Aspose.Cells | Stacked bar chart as progress indicator Aspose.Cells | C# code for dynamic progress bars in Excel with Aspose
// Developer Intent: Generate a stacked bar chart that acts as a progress bar by linking percentage values from a worksheet column to the chart’s visible series.
// Use Cases: Project status reports that display each task’s completion as a visual progress bar. | Automated dashboards that render multiple progress indicators without manual chart editing. | Weekly update sheets that export task names and dynamically drawn progress bars directly from data.
// AI Prompts: Write C# code using Aspose.Cells to bind a numeric column to a stacked bar chart series and assign category labels from another column. | Explain how to customize bar colors, add data labels, and format percentages for a progress bar chart created with Aspose.Cells. | Show how to keep raw fractional values in the worksheet while displaying them as percentages on the chart.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to generate an Excel workbook with Aspose.Cells, fill column A with task names and column B with fractional progress values, add a stacked bar chart, bind B2:B5 as the series data and A2:A5 as the category axis, configure the series to display as a progress bar, and save the file as ProgressBarChart.xlsx.
class ProgressBarChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: task names in column A and progress percentages in column B
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Progress");
        string[] tasks = { "Design", "Development", "Testing", "Deployment" };
        double[] progresses = { 0.25, 0.5, 0.75, 0.9 }; // values as fractions (25%, 50%, etc.)

        for (int i = 0; i < tasks.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(tasks[i]);      // A column (category)
            sheet.Cells[i + 2, 1].PutValue(progresses[i]); // B column (values)
        }

        // Add a bar chart that will act as a progress bar
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the progress percentages (B2:B5) to the visible series of the chart
        // Add the series data range (vertical) and set the category (task names)
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Ensure the series is displayed (not filtered) and use a stacked bar for visual effect
        chart.NSeries[0].IsFiltered = false;
        chart.NSeries[0].Type = ChartType.BarStacked;

        // Save the workbook to a file
        workbook.Save("ProgressBarChart.xlsx");
    }
}
