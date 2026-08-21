// Title: Insert a column chart into a 3‑row × 5‑column range using Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, adds sample data, and places a column chart whose top‑left corner starts at cell B7. The chart area is limited to exactly three rows and five columns, then the series is bound to A1:B5 and the file is saved as ChartInRange.xlsx.
// Keywords: Aspose.Cells | C# | .NET | insert chart | specific cell range | column chart positioning | 3 rows 5 columns | Excel automation | chart area size | B7 chart placement
// Common Searches: Aspose.Cells add chart to a defined cell range C# | set chart size to 3 rows and 5 columns Aspose.Cells | place column chart at B7 using Aspose.Cells for .NET | how to limit chart area to specific cells in Excel with Aspose | C# code to insert chart into exact cell block
// Developer Intent: Place a chart in a workbook so it exactly fills a 3‑row by 5‑column block.
// Use Cases: Generate a sales dashboard where each chart must occupy a fixed 3 × 5 cell area next to the data table. | Create a printable financial report template that reserves a specific cell block for charts to maintain consistent layout. | Automate monthly KPI sheets that embed column charts within predefined cell boundaries for easy page‑break handling.
// AI Prompts: Show how to insert a pie chart into cells C10:G12 with Aspose.Cells for .NET (C#). | Explain how to compute chart range dynamically based on the number of data rows in Aspose.Cells. | Provide C# code to resize an existing chart while keeping it inside a 4‑row × 6‑column area after adding a new series.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a new workbook, adds sample data, and places a column chart whose top‑left corner starts at cell B7. The chart area is limited to exactly three rows and five columns, then the series is bound to A1:B5 and the file is saved as ChartInRange.xlsx.
class InsertChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Item " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Define the top‑left corner of the chart (zero‑based indices)
        int topRow = 6;        // Row 7
        int leftColumn = 1;    // Column B

        // Chart must occupy exactly 3 rows and 5 columns
        int bottomRow = topRow + 2;   // 3 rows total
        int rightColumn = leftColumn + 4; // 5 columns total

        // Add a column chart to the specified range
        int chartIndex = sheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("=Sheet1!$A$1:$B$5", true);

        // Save the workbook
        workbook.Save("ChartInRange.xlsx", SaveFormat.Xlsx);
    }
}
