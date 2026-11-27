// Title: Insert a column chart into a 3‑row by 5‑column range in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# place a column chart in a defined cell range | set chart position by top row and left column indices with Aspose.Cells | create a chart that occupies exactly three rows and five columns in Excel using Aspose.Cells | C# Aspose.Cells specify chart bounds using row and column numbers | how to define chart area size in Aspose.Cells workbook
// Tags: Aspose.Cells insert column chart in worksheet range | C# define chart bounds with topRow leftColumn indices | Aspose.Cells chart area size rows columns | Excel chart placement using Aspose.Cells .NET | C# create chart with specific cell coordinates Aspose | Aspose.Cells chart positioning using row column indices | Excel chart size specification Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds sample data, calculates a 3‑row by 5‑column area starting at row 5 column 2, inserts a column chart into that range, assigns data series and a title, and saves the file as ChartInRange.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Define the cell range where the chart should be placed.
        // The chart must occupy exactly 3 rows and 5 columns.
        // Aspose.Cells uses zero‑based indices.
        int topRow = 5;          // first row of the chart area
        int leftColumn = 2;      // first column of the chart area
        int bottomRow = topRow + 2;   // 3 rows total (topRow, topRow+1, topRow+2)
        int rightColumn = leftColumn + 4; // 5 columns total (leftColumn .. leftColumn+4)

        // Add a column chart positioned in the defined range
        int chartIndex = sheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        // Values: B2:B4, Categories: A2:A4
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Optional: set a chart title
        chart.Title.Text = "Quarterly Sales";

        // Save the workbook
        try
        {
            workbook.Save("ChartInRange.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully as ChartInRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error saving workbook: " + ex.Message);
        }
    }
}
