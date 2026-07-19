// Title: Aspose.Cells for .NET – C# example: Line chart with moving‑average series from worksheet formula
// Description: This C# sample builds a workbook, writes index and sample values, inserts a 3‑period moving average using the AVERAGE function, evaluates all formulas, creates a line chart that reads the calculated averages as its series, maps the X‑axis to the index column, refreshes the chart, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | moving average chart | worksheet formula chart source | line chart Aspose.Cells | calculate moving average | chart series from formula | Excel chart automation | Aspose.Cells example
// Common Searches: Aspose.Cells chart series from formula | C# moving average line chart Aspose.Cells | How to use worksheet formulas as chart data in .NET | Create dynamic chart with calculated values Aspose.Cells | Generate moving average chart programmatically
// Developer Intent: Create a line chart whose data series is derived from a moving‑average formula in the worksheet.
// Use Cases: Show sales trend with a 3‑period moving average that updates automatically when source data changes. | Produce a chart ready for export to PNG or PDF after calling chart.Calculate(). | Integrate dynamic trend lines into financial reports generated with Aspose.Cells.
// AI Prompts: Modify the code to calculate a 5‑period moving average. | Add code to export the chart as PNG and PDF files. | Include the original values as a second series on the same chart. | Parameterize the moving‑average period via a variable or user input. | Explain how to bind the chart to a named range instead of explicit cell addresses.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMovingAverageChart
{
    // This C# sample builds a workbook, writes index and sample values, inserts a 3‑period moving average using the AVERAGE function, evaluates all formulas, creates a line chart that reads the calculated averages as its series, maps the X‑axis to the index column, refreshes the chart, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Populate sample data ----------
            // Category (X‑axis) in column A
            // Original values in column B
            // Moving average (period = 3) will be placed in column C
            int dataCount = 12;
            int movingAvgPeriod = 3;

            // Header row
            cells["A1"].PutValue("Index");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("MovingAvg");

            // Fill columns A and B with sample data
            for (int i = 0; i < dataCount; i++)
            {
                cells[i + 1, 0].PutValue(i + 1);                 // A column (Index)
                cells[i + 1, 1].PutValue(10 + i * 5);           // B column (Value)
            }

            // ---------- Insert moving‑average formula in column C ----------
            // For rows where a full period is not available, leave the cell blank
            for (int row = movingAvgPeriod; row <= dataCount; row++)
            {
                // Formula: =AVERAGE(B{row‑period+1}:B{row})
                string formula = $"=AVERAGE(B{row - movingAvgPeriod + 1}:B{row})";
                cells[row, 2].Formula = formula;
            }

            // Calculate all formulas so that column C contains the moving‑average values
            workbook.CalculateFormula();

            // ---------- Add a line chart ----------
            // Position the chart (row, column, height, width) in the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series to use the moving‑average values (C2:C{dataCount})
            // Category (X‑axis) data comes from the Index column (A2:A{dataCount})
            chart.NSeries.Add($"C2:C{dataCount}", true);
            chart.NSeries.CategoryData = $"A2:A{dataCount}";

            // Optional: give the series a name
            chart.NSeries[0].Name = "3‑Period Moving Average";

            // Recalculate the chart (important if you plan to save as an image later)
            chart.Calculate();

            // ---------- Save the workbook ----------
            workbook.Save("MovingAverageChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
