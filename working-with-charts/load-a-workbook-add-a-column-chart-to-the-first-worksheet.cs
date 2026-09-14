// Title: Load an existing Excel workbook and insert a column chart into the first worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a .xlsx file, adds sample data if needed, creates a Column chart positioned at rows 5‑15 and columns 0‑5, binds it to range A1:B4, and saves the workbook. | Show how to use Aspose.Cells.ChartType.Column and the NSeries.Add method to programmatically generate a column chart from a data range in a loaded workbook. | Provide a step‑by‑step snippet that loads a workbook, adds a column chart to the first sheet, sets its data source, and exports the result as a new Excel file.
// Common Searches: Aspose.Cells C# how to add a column chart to an existing worksheet | programmatically create a column chart from range A1:B4 in .NET | load Excel file and insert chart with Aspose.Cells example | C# sample code for adding chart to first sheet and saving workbook
// Tags: add column chart Aspose.Cells .NET | bind chart to range A1:B4 Aspose.Cells | insert chart into first worksheet programmatically | save workbook with new chart Aspose.Cells | generate sample data for chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Loads input.xlsx, optionally writes sample data, adds a column chart covering cells A1:B4 on the first worksheet, and saves the result as output.xlsx.
class AddColumnChartExample
{
    static void Main()
    {
        // Path to the existing workbook to load
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // OPTIONAL: add sample data if the workbook does not already contain it
        // This ensures the chart has a data source to display
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart (A1:B4) and indicate that data is plotted by column
        chart.NSeries.Add("A1:B4", true);

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
