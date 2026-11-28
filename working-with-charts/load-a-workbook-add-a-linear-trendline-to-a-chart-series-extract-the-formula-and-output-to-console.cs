// Title: Create a scatter chart, add a linear trendline, and read its equation with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add linear trendline to scatter chart and get equation | How to retrieve trendline formula and R-squared from an Aspose.Cells chart series | Aspose.Cells .NET calculate workbook to display trendline equation in data label | Extract trendline DataLabels.Text after adding linear trendline in Aspose.Cells | Save workbook with trendline equation using Aspose.Cells C# example
// Tags: add linear trendline Aspose.Cells scatter chart | extract trendline equation Aspose.Cells C# | display r-squared Aspose.Cells chart series | calculate workbook formulas Aspose.Cells .NET | save workbook with trendline Aspose.Cells xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TrendlineFormulaDemo
{
    // The sample program creates a new workbook, populates columns A and B with X‑Y data, inserts a scatter chart, binds the Y series and X values, adds a linear trendline with equation and R‑squared display enabled, forces formula calculation, reads the generated equation from the trendline's DataLabels.Text, prints it to the console, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data (X values in column A, Y values in column B)
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                for (int i = 2; i <= 10; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);               // X = 1,2,3,...
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 2 + 5);   // Y = 7,9,11,...
                }

                // 3. Add a scatter chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Scatter, 12, 0, 28, 15);
                Chart chart = sheet.Charts[chartIndex];

                // 4. Add the data series (Y values) and set X values
                chart.NSeries.Add("B2:B10", true);
                chart.NSeries[0].XValues = "A2:A10";

                // 5. Add a linear trendline to the first series
                int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
                Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

                // Optional: give the trendline a custom name
                trendline.Name = "Linear Trend";

                // 6. Request the equation and R‑squared value to be displayed
                trendline.DisplayEquation = true;
                trendline.DisplayRSquared = true;

                // 7. Force calculation so that the trendline data (equation) is generated
                workbook.CalculateFormula();

                // 8. Extract the formula (equation) from the trendline.
                // The equation appears in the trendline's DataLabels text.
                string equation = trendline.DataLabels.Text;

                // 9. Output the result to the console
                Console.WriteLine("Trendline Equation (including R‑squared if displayed):");
                Console.WriteLine(equation);

                // 10. Save the workbook (optional, demonstrates that the chart is correct)
                string outFile = "TrendlineFormulaDemo.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
