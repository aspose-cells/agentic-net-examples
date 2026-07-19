// Title: Create an Excel workbook with a column chart using Aspose.Cells for .NET (C#)
// Description: This example shows how to create a new Workbook, add sample data to the first Worksheet, insert a Column chart positioned from row 5 col 2 to row 25 col 11, bind it to the range A1:B11, and save the file as ColumnChartOutput.xlsx.
// Keywords: Aspose.Cells column chart C# | add chart to Excel workbook .NET | generate Excel file with chart Aspose | C# column chart example | Aspose.Cells chart API
// Common Searches: how to insert a column chart with Aspose.Cells | Aspose.Cells C# sample for creating charts | create and save Excel chart programmatically .NET | column chart example Aspose.Cells
// Developer Intent: Programmatically generate an Excel file that includes a column chart based on populated data.
// Use Cases: Automate monthly sales reports with visual column charts. | Build a performance dashboard that adds data and visualizes it as a column chart. | Create a reusable template that inserts a column chart for any data range before exporting.
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, fill A1:B15 with headers and numbers, add a clustered column chart covering rows 5‑25 and columns 2‑11, bind it to the data range, and save as .xlsx. | Explain step‑by‑step how to add a column chart in Aspose.Cells, covering chart type selection, positioning, and data series linking. | Modify the sample to use a stacked column chart, set a custom chart title, and add axis labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to create a new Workbook, add sample data to the first Worksheet, insert a Column chart positioned from row 5 col 2 to row 25 col 11, bind it to the range A1:B11, and save the file as ColumnChartOutput.xlsx.
public class CreateWorkbookWithColumnChart
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 10; i++)
        {
            sheet.Cells[$"A{i + 1}"].PutValue($"Cat {i}");
            sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
        }

        // Insert a column chart into the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 25, 11);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("=Sheet1!$A$1:$B$11", true);

        // Save the workbook to a file
        workbook.Save("ColumnChartOutput.xlsx", SaveFormat.Xlsx);
    }
}
