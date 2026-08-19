// Title: Hide Gridlines on a Pie Chart with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, add data, create a pie chart, and turn off both major and minor gridlines on the category and value axes using Aspose.Cells for .NET, delivering a chart with no visible gridlines.
// Keywords: Aspose.Cells C# hide chart gridlines | pie chart gridlines Aspose.Cells | remove gridlines Aspose.Cells .NET | chart axis gridlines visibility | disable major minor gridlines | Aspose.Cells chart styling
// Common Searches: Aspose.Cells hide pie chart gridlines C# | C# remove chart gridlines Aspose.Cells | set chart gridlines invisible Aspose.Cells | pie chart without gridlines Aspose.Cells .NET | how to turn off chart axes gridlines Aspose.Cells
// Developer Intent: Disable all gridlines on a pie chart to achieve a cleaner visual output.
// Use Cases: Generate a sales‑distribution pie chart for a financial report and suppress gridlines before exporting to Excel. | Automate dashboard creation where every pie chart must match corporate style guidelines that forbid axis gridlines. | Build a multi‑sheet workbook with numerous pie charts, ensuring each chart displays without gridlines for improved readability.
// AI Prompts: Write C# code with Aspose.Cells that creates a pie chart and makes both major and minor gridlines on its axes invisible. | Explain step‑by‑step how to guarantee a pie chart in Aspose.Cells shows no gridlines, including which axis properties to modify. | Provide a method to programmatically verify that gridlines are hidden on an Aspose.Cells chart before saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, add data, create a pie chart, and turn off both major and minor gridlines on the category and value axes using Aspose.Cells for .NET, delivering a chart with no visible gridlines.
class RemovePieChartGridlines
{
    static void Main(string[] args)
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
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a pie chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide all gridlines (major and minor) on both axes
        // Pie charts may not display axes, but setting visibility to false guarantees they are removed
        chart.CategoryAxis.MajorGridLines.IsVisible = false;
        chart.CategoryAxis.MinorGridLines.IsVisible = false;
        chart.ValueAxis.MajorGridLines.IsVisible = false;
        chart.ValueAxis.MinorGridLines.IsVisible = false;

        // Save the workbook with the cleaned-up pie chart
        string outputPath = "PieChart_NoGridlines.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
