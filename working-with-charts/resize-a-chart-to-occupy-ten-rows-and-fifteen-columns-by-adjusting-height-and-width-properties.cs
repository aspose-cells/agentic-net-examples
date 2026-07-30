// Title: C# – Resize an Aspose.Cells chart to 10 rows × 15 columns using Chart.Move
// Description: Creates a workbook, adds a column chart with sample data, then calls Chart.Move with calculated top‑row, left‑column, bottom‑row and right‑column indices so the chart covers exactly ten rows and fifteen columns before saving the file.
// Keywords: Aspose.Cells chart resize C# | Chart.Move method | set chart size rows columns | Excel chart dimensions .NET | adjust chart height width Aspose | C# Aspose.Cells example
// Common Searches: Aspose.Cells resize chart to specific cells | C# chart.Move rows columns example | How to set chart width and height in Aspose.Cells | Resize Excel chart programmatically C# | Chart.Move parameters explained
// Developer Intent: Programmatically set a chart’s bounds so it spans a defined number of worksheet rows and columns.
// Use Cases: Place a generated chart into a fixed grid area for consistent report layouts. | Adapt chart size dynamically based on the amount of data displayed while keeping a uniform cell footprint. | Ensure charts align with other worksheet elements such as tables or images across multiple generated workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to move a chart so it occupies rows 5‑14 and columns 2‑16. | Explain how the four parameters of Chart.Move correspond to chart height and width in cell units. | Show an alternative way to resize a chart by setting its Height and Width properties instead of using Move.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartResize
{
    // Creates a workbook, adds a column chart with sample data, then calls Chart.Move with calculated top‑row, left‑column, bottom‑row and right‑column indices so the chart covers exactly ten rows and fifteen columns before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart (optional, just to have a visible chart)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart. Initial position is arbitrary; it will be resized later.
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 0, 0, 5, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Resize the chart to occupy 10 rows and 15 columns.
            // Choose a starting cell (top-left corner) at row 5, column 2 (zero‑based indices).
            int topRow = 5;          // Upper left row index
            int leftColumn = 2;      // Upper left column index
            int bottomRow = topRow + 9;   // 10 rows total (0‑based inclusive)
            int rightColumn = leftColumn + 14; // 15 columns total

            // Use the Move method to set the new bounds.
            chart.Move(topRow, leftColumn, bottomRow, rightColumn);

            // Save the workbook
            workbook.Save("ResizedChart.xlsx");
        }
    }
}
