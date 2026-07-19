// Title: Bind an Aspose.Cells chart series to a named range (SalesData) in C# for dynamic updates
// Description: C# example that creates a workbook, defines a named range called SalesData, adds a column chart, links the series to the named range with NSeries.Add, optionally sets month labels as CategoryData, and saves the file so the chart refreshes when the range changes.
// Keywords: Aspose.Cells | C# | .NET | named range | chart series | dynamic chart | NSeries.Add | column chart | Excel automation | programmatic chart
// Common Searches: Aspose.Cells bind chart series to named range C# | How to use named ranges with charts in Aspose.Cells .NET | Dynamic chart updates using Aspose.Cells named range | C# Aspose.Cells NSeries.Add example | Set chart category axis from cells Aspose.Cells
// Developer Intent: Create a chart whose data source is a named range so the chart updates automatically when the range values change.
// Use Cases: Define a SalesData named range and link it to a column chart, allowing new sales rows to appear in the chart without rewriting formulas. | Assign month names as CategoryData to display readable axis labels alongside the dynamic series. | Reuse the same named range across multiple charts in a workbook to keep all visualizations synchronized with a single data source.
// AI Prompts: Generate C# code using Aspose.Cells that creates a named range for chart data and binds it to a column chart series. | Show how to modify an existing Aspose.Cells chart to reference a named range for both series values and category labels. | Explain how to expand the SalesData named range programmatically when new rows are added so the chart grows automatically.

using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, defines a named range called SalesData, adds a column chart, links the series to the named range with NSeries.Add, optionally sets month labels as CategoryData, and saves the file so the chart refreshes when the range changes.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (Month in column A, Sales in column B)
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
        int[] sales = { 120, 150, 130, 170, 160 };
        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(months[i]);   // A2:A6
            sheet.Cells[i + 1, 1].PutValue(sales[i]);   // B2:B6
        }

        // Create a named range called "SalesData" that refers to the sales values (B2:B6)
        int nameIndex = workbook.Worksheets.Names.Add("SalesData");
        workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$B$2:$B${months.Length + 1}";

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Link the chart series to the named range "SalesData"
        chart.NSeries.Add("SalesData", true);

        // (Optional) Link the category axis to the month names
        chart.NSeries.CategoryData = $"={sheet.Name}!$A$2:$A${months.Length + 1}";

        // Save the workbook
        workbook.Save("ChartWithNamedRange.xlsx");
    }
}
