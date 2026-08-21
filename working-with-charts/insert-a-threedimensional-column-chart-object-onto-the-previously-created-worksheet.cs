// Title: Add a 3‑D Column Chart to an Excel Worksheet using Aspose.Cells for C#
// Description: Creates a new Workbook, writes sample sales data to A1:B4, inserts a three‑dimensional column chart (ChartType.Column3D) positioned from rows 5‑20 and columns 0‑8, binds the chart to the data range, and saves the file as ThreeDColumnChart.xlsx.
// Keywords: Aspose.Cells 3D column chart C# | ChartType.Column3D Aspose | add 3D chart Aspose.Cells | set chart data range Aspose | save workbook with chart Aspose | C# Excel chart generation | Aspose.Cells chart positioning
// Common Searches: how to insert a 3d column chart with Aspose.Cells .NET | Aspose.Cells C# create 3‑D column chart | set data range for 3D chart Aspose.Cells | save Excel file containing a 3D chart using Aspose | Aspose.Cells chart placement rows columns
// Developer Intent: Generate a three‑dimensional column chart on a worksheet and persist the workbook as an .xlsx file.
// Use Cases: Visualize quarterly sales figures in a 3‑D column chart for management reports. | Provide a reusable routine that adds a 3‑D column chart to any worksheet given a data range and position. | Export Excel dashboards with embedded 3‑D charts for stakeholder presentations.
// AI Prompts: Write a C# function that accepts a Worksheet, a data‑range string, and chart coordinates, then adds a 3‑D column chart with Aspose.Cells. | Show error‑handling code for creating a 3‑D column chart and saving the workbook, including logging of exceptions. | Demonstrate how to customize the title, axis labels, and series colors of a 3‑D column chart after it is added.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new Workbook, writes sample sales data to A1:B4, inserts a three‑dimensional column chart (ChartType.Column3D) positioned from rows 5‑20 and columns 0‑8, binds the chart to the data range, and saves the file as ThreeDColumnChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Q1");
        worksheet.Cells["A3"].PutValue("Q2");
        worksheet.Cells["A4"].PutValue("Q3");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(1500);
        worksheet.Cells["B4"].PutValue(1800);

        // Add a three‑dimensional column chart to the worksheet
        // Parameters: chart type, top row, left column, bottom row, right column
        int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        chart.SetChartDataRange("A1:B4", true);

        // The chart is now a 3D column chart; Is3D property will return true
        // Console.WriteLine($"Is3D: {chart.Is3D}");

        // Save the workbook with the chart
        workbook.Save("ThreeDColumnChart.xlsx", SaveFormat.Xlsx);
    }
}
