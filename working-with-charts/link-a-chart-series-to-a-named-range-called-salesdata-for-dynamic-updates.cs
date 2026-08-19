// Title: Link a chart series to a named range for dynamic updates with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a named range called "SalesData", add a column chart, and bind the chart series to that range using Aspose.Cells. The chart automatically reflects any changes made to the named range, enabling live data visualisation without rebuilding the series.
// Keywords: Aspose.Cells chart named range | C# dynamic chart data | bind chart series to named range | Aspose.Cells NSeries Add | named range chart update .NET
// Common Searches: Aspose.Cells bind chart series to named range C# | dynamic chart data using named ranges Aspose.Cells | C# add named range for chart Aspose.Cells | update chart automatically when range changes Aspose.Cells | how to link chart to named range Aspose.Cells .NET
// Developer Intent: Connect a chart series to a predefined named range so the chart refreshes automatically when the range values are modified.
// Use Cases: Create a reusable named range for sales figures and attach it to a column chart for real‑time updates. | Modify data in the "SalesData" range later and have the existing chart reflect the new values instantly. | Share a single named range across multiple charts to maintain consistent data sources throughout a workbook.
// AI Prompts: Generate C# code with Aspose.Cells that defines a named range and links a chart series to it for automatic refresh. | Show how to change values in a named range after a chart is created and ensure the chart updates without extra code. | Explain how to set X‑axis (category) data for a chart using a named range in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a named range called "SalesData", add a column chart, and bind the chart series to that range using Aspose.Cells. The chart automatically reflects any changes made to the named range, enabling live data visualisation without rebuilding the series.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (Month and Sales)
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
        int[] sales = { 120, 150, 130, 170, 160 };

        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
            sheet.Cells[i + 1, 1].PutValue(sales[i]);    // Column B
        }

        // Define a named range called "SalesData" that refers to the sales values (B2:B6)
        int nameIndex = workbook.Worksheets.Names.Add("SalesData");
        // Note: RefersTo must include the sheet name and absolute addresses
        workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$B$2:$B${months.Length + 1}";

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Link the chart series to the named range "SalesData"
        // The Add method can accept the name directly (without '=') for dynamic updates
        chart.NSeries.Add("SalesData", true);

        // Optionally set the category (X‑axis) data to the month names
        chart.NSeries.CategoryData = $"={sheet.Name}!$A$2:$A${months.Length + 1}";

        // Save the workbook
        workbook.Save("ChartWithNamedRange.xlsx");
    }
}
