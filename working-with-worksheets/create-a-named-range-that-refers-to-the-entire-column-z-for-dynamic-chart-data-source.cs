// Title: Create a workbook‑level named range for column Z and bind it to a dynamic chart using Aspose.Cells for .NET (C#)
// Description: The example creates a new workbook, fills column Z with sample values, adds a workbook‑level named range that references the entire column Z ("=Sheet1!$Z:$Z"), creates a column chart, sets the chart's data source to the named range, and saves the file as NamedRangeColumnZ.xlsx.
// Keywords: Aspose.Cells | named range | column Z | dynamic chart | C# | .NET | chart data source | entire column reference | SetChartDataRange | Workbook.Names.Add
// Common Searches: Aspose.Cells named range entire column | C# set chart data source named range Aspose.Cells | reference whole column in Aspose.Cells | dynamic chart from column Z Aspose.Cells | add workbook level named range Aspose.Cells .NET
// Developer Intent: Define a workbook‑level named range that points to the full column Z and use it as the chart data source so the visualization automatically expands as new rows are added.
// Use Cases: Create dashboards where the chart updates automatically when more rows are added to column Z. | Reuse the same ColumnZData range across multiple charts or worksheets for consistent visualizations. | Change the column reference in one place to update all linked charts without modifying each chart individually. | Generate reports that require a flexible data range without hard‑coding row limits.
// AI Prompts: Generate C# code with Aspose.Cells that defines a named range for the entire column Z and assigns it to a chart data series. | Explain how to modify the named range to point to a different column or a specific row interval in Aspose.Cells. | Show steps to replace an existing chart's data source with a new named range for dynamic data in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills column Z with sample values, adds a workbook‑level named range that references the entire column Z ("=Sheet1!$Z:$Z"), creates a column chart, sets the chart's data source to the named range, and saves the file as NamedRangeColumnZ.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate some sample data in column Z (index 25, zero‑based)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 25].PutValue(i + 1);
        }

        // Add a named range that refers to the entire column Z
        int nameIdx = workbook.Worksheets.Names.Add("ColumnZData");
        Name columnZName = workbook.Worksheets.Names[nameIdx];
        // The RefersTo formula must start with '=' and use absolute column reference
        columnZName.RefersTo = "=Sheet1!$Z:$Z";

        // Create a chart and set its data source to the named range
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIdx];
        // Use the named range; true indicates plotting by column (vertical)
        chart.SetChartDataRange("ColumnZData", true);

        // Save the workbook
        workbook.Save("NamedRangeColumnZ.xlsx");
    }
}
