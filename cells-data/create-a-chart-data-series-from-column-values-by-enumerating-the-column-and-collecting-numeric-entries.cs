// Title: Create a column chart from numeric cells in a worksheet column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create column chart using only numeric cells from a worksheet column | How to filter out text values when building a chart series with Aspose.Cells .NET | Extract numeric values from column A and use them as a data range for a chart in C# | Generate dynamic chart series from a list of doubles in an Excel file with Aspose.Cells
// Tags: filter numeric cells Aspose.Cells | dynamic chart series range C# | column chart from worksheet data Aspose.Cells | extract numeric values column Aspose.Cells | build chart data range programmatically .NET | create chart series from list of doubles Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.IO;

// The example creates a workbook, fills column A with mixed data, iterates the column to collect only numeric values, writes those numbers to column B, adds a column chart whose series references the B column range, sets a chart title, and saves the file as ChartFromColumn.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data in column A (index 0) – includes numeric and non‑numeric entries
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Text");
            sheet.Cells["A4"].PutValue(20.5);
            sheet.Cells["A5"].PutValue(15);
            sheet.Cells["A6"].PutValue(null);
            sheet.Cells["A7"].PutValue(30);

            // Enumerate column A and collect only numeric values
            List<double> numericValues = new List<double>();
            int lastRow = sheet.Cells.MaxDataRow; // last row with data in the sheet
            for (int row = 1; row <= lastRow; row++) // start at row 1 to skip header
            {
                Cell cell = sheet.Cells[row, 0]; // column A
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericValues.Add(cell.DoubleValue);
                }
            }

            // Write numeric values to column B so they can be used as a data range for the chart
            int dataStartRow = 1; // zero‑based index (row 2 in Excel)
            for (int i = 0; i < numericValues.Count; i++)
            {
                sheet.Cells[dataStartRow + i, 1].PutValue(numericValues[i]); // column B (index 1)
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series (e.g., B2:B{n})
            string dataRange = $"B{dataStartRow + 1}:B{dataStartRow + numericValues.Count}";
            chart.NSeries.Add(dataRange, true);
            chart.NSeries[0].Name = "Series1";

            // Set a title for the chart (optional)
            chart.Title.Text = "Numeric Values from Column A";

            // Determine output path and ensure directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartFromColumn.xlsx");
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
