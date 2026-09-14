// Title: Create a column chart and freeze the first four rows in an Aspose.Cells worksheet using C#
// AI Prompts: Write C# code that fills cells A1:B4, creates a column chart from the values, applies FreezePanes to rows 1‑4, and saves the workbook. | Generate an Aspose.Cells .NET example that adds sample data, inserts a column chart, locks the first four rows with FreezePanes, and outputs ChartWithFrozenRows.xlsx. | Provide a snippet using Aspose.Cells for C# to build a column chart, freeze the header rows, and write the file to disk.
// Common Searches: Aspose.Cells C# freeze top rows after creating a column chart | how to use FreezePanes with chart data in Aspose.Cells .NET | example of adding a column chart and keeping data rows visible in an Excel file using Aspose.Cells
// Tags: column chart creation Aspose.Cells C# | FreezePanes rows Aspose.Cells .xlsx | chart NSeries data range Aspose.Cells | save workbook with frozen rows Aspose.Cells | add chart and lock rows Aspose.Cells example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// // Creates a new workbook, writes sample data, adds a column chart based on the values, freezes the first four rows to keep the data visible, and saves the file as ChartWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a column chart below the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true); // Values
            // Category data can be inferred; explicit assignment omitted for compatibility
            chart.Title.Text = "Sample Column Chart";

            // Freeze the first four rows (rows 1‑4)
            // FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(4, 0, 4, 0);

            // Save the workbook
            workbook.Save("ChartWithFrozenRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
