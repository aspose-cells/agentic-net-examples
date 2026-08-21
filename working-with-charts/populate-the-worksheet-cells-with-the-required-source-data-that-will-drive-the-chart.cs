// Title: Create a Column Chart from Populated Cells with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to fill a worksheet with header and data rows, add a vertical column chart that references range A1:C5, set a custom title, position the chart, and save the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart example | populate Excel cells Aspose | column chart from range Aspose.Cells | set chart title .NET | chart positioning Aspose.Cells | save workbook as xlsx | Excel automation C# | Aspose.Cells add chart with data range | vertical column chart .NET | programmatic chart generation
// Common Searches: Aspose.Cells create column chart from cells | C# add chart with data range A1:C5 | how to set chart title in Aspose.Cells | position chart in worksheet Aspose.Cells | populate worksheet then chart Aspose .NET
// Developer Intent: Populate worksheet data and generate a column chart programmatically with Aspose.Cells in C#.
// Use Cases: Automated sales reporting: write product categories and two sales series to Excel, then visualize them with a column chart. | Dynamic dashboard generation: update data arrays in code, refresh the chart source range, and embed the chart at a specific location on the sheet. | Export of analytical results: create an XLSX file that includes a titled column chart for stakeholder presentations.
// AI Prompts: Generate C# code using Aspose.Cells that writes category labels and two numeric series to cells A1:C5, adds a vertical column chart covering that range, sets the title "Sample Column Chart", positions the chart between rows 5‑20 and columns 1‑10, and saves the file as ChartWithData.xlsx. | Provide an Aspose.Cells .NET snippet that reads string and integer arrays, populates the worksheet, creates a column chart with a custom title, defines its placement, and exports the workbook to XLSX.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataExample
{
    // Demonstrates how to fill a worksheet with header and data rows, add a vertical column chart that references range A1:C5, set a custom title, position the chart, and save the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the cells collection
            Cells cells = sheet.Cells;

            // Populate header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Series1");
            cells["C1"].PutValue("Series2");

            // Populate sample data (rows 2 to 5)
            string[] categories = { "A", "B", "C", "D" };
            int[] series1 = { 10, 30, 50, 70 };
            int[] series2 = { 20, 40, 60, 80 };

            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // Excel rows are 1‑based; start at row 2
                cells[$"A{row}"].PutValue(categories[i]);
                cells[$"B{row}"].PutValue(series1[i]);
                cells[$"C{row}"].PutValue(series2[i]);
            }

            // Add a column chart using the overload that accepts data range and positioning
            // Data range: A1:C5 (including headers)
            // isVertical = true (plot series by column)
            // Position: topRow=5, leftColumn=1, rightRow=20, bottomColumn=10
            int chartIndex = sheet.Charts.Add(ChartType.Column, "A1:C5", true, 5, 1, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Optionally set chart title
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook
            workbook.Save("ChartWithData.xlsx", SaveFormat.Xlsx);
        }
    }
}
