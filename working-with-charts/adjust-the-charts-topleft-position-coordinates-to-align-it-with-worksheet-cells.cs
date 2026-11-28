// Title: Set a chart’s top‑left and bottom‑right cell coordinates to align it with worksheet cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# place chart at specific cell range B2:H20 | how to define chart bounds with row and column indices in Aspose.Cells | C# Aspose.Cells set chart top left and bottom right cells programmatically | position column chart inside worksheet cells using Aspose.Cells API | align Aspose.Cells chart to worksheet cells by specifying topRow leftColumn bottomRow rightColumn
// Tags: chart positioning using cell indices Aspose.Cells | set chart bounds with topRow leftColumn bottomRow rightColumn | align chart to worksheet cell range .NET | Aspose.Cells column chart placement by cells | define chart area B2 H20 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills columns A and B with sample data, defines a cell range from B2 to H20 using row and column indices, adds a column chart positioned within that range, sets the chart's data source, assigns a title, prints the placement coordinates, and saves the file as ChartAlignedWithCells.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Data";

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            for (int i = 1; i <= 10; i++)
            {
                worksheet.Cells[i + 1, 0].PutValue("Cat " + i);   // Column A
                worksheet.Cells[i + 1, 1].PutValue(i * 5);       // Column B
            }

            // Define the cell range that will bound the chart.
            // Top‑left cell B2 (row index 1, column index 1)
            // Bottom‑right cell H20 (row index 19, column index 7)
            int topRow = 1;        // B2 row
            int leftColumn = 1;    // B2 column
            int bottomRow = 19;    // H20 row
            int rightColumn = 7;   // H20 column

            // Add a column chart and position it using the cell coordinates above
            int chartIndex = worksheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("=Data!$A$1:$B$11", true);
            chart.Title.Text = "Sample Column Chart";

            // Optional: display the actual placement values
            Console.WriteLine($"Chart placed from row {topRow} to {bottomRow}, column {leftColumn} to {rightColumn}");

            // Save the workbook
            string outputPath = "ChartAlignedWithCells.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
