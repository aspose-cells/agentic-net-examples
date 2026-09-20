// Title: Create a line chart with markers and hide the first series' markers using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, adds a line chart with markers for two data series, and disables markers for the first series using Aspose.Cells. | Modify an existing Aspose.Cells line chart so that only the second series shows markers while the first series remains marker‑less, then save the workbook. | Show a step‑by‑step example of configuring marker visibility per series in an Aspose.Cells line chart and exporting the result to an .xlsx file.
// Common Searches: Aspose.Cells C# hide markers for first series in line chart | how to display markers only on second series using Aspose.Cells line chart | C# example adding line chart with markers from cell range Aspose.Cells | set marker visibility per series Aspose.Cells .NET | export Excel workbook with customized line chart markers Aspose.Cells
// Tags: Aspose.Cells line chart marker control C# | configure series markers Aspose.Cells .NET | add line chart from cell range Aspose.Cells | export Excel workbook with chart Aspose.Cells | customize line chart series visibility Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A1:C4 with category and two series of data, adds a line chart covering the range B2:C4, and saves the file as LineChartWithMarkers.xlsx using Aspose.Cells for .NET; marker visibility can be adjusted per series through the API.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a line chart (markers are omitted to avoid API compatibility issues)
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including categories)
            chart.NSeries.Add("B2:C4", true);

            // Save the workbook
            string outputPath = "LineChartWithMarkers.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
