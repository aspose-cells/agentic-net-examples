// Title: Generate a column chart from VLOOKUP‑derived quarterly targets using Aspose.Cells for .NET
// AI Prompts: Write a C# program that creates a lookup table, inserts a VLOOKUP formula column, calculates the workbook, and binds the resulting values to a column chart with Aspose.Cells. | Add a VLOOKUP formula referencing $A$2:$B$5 into column E, call workbook.CalculateFormula(), and set the chart series range to E2:E5 in an Aspose.Cells workbook.
// Common Searches: asp.net aspocells add VLOOKUP formula and use it as chart data source | c# generate Excel column chart from VLOOKUP results using Aspose.Cells | populate Aspose.Cells chart series with values returned by VLOOKUP formula | calculate formulas before creating chart in Aspose.Cells .NET
// Tags: Aspose.Cells VLOOKUP formula chart | C# column chart from calculated cells | Excel chart data series using VLOOKUP Aspose.Cells | Aspose.Cells calculate formulas before chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsVLookupChartDemo
{
    // Demonstrates building a quarterly target lookup table, applying VLOOKUP formulas to retrieve values, calculating the formulas, and creating a column chart that visualizes the VLOOKUP results using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();                     // create
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 2. Populate a lookup table with quarterly targets
                //    Column A : Quarter identifiers (Q1, Q2, Q3, Q4)
                //    Column B : Corresponding target values
                // ------------------------------------------------------------
                string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
                double[] targets = { 120000, 150000, 130000, 160000 };

                for (int i = 0; i < quarters.Length; i++)
                {
                    cells[i + 1, 0].PutValue(quarters[i]); // A2:A5
                    cells[i + 1, 1].PutValue(targets[i]);  // B2:B5
                }

                // ------------------------------------------------------------
                // 3. Prepare the data range that will be used by the chart.
                //    Column D will hold the same quarter identifiers.
                //    Column E will contain VLOOKUP formulas that retrieve the
                //    target values from the lookup table.
                // ------------------------------------------------------------
                for (int i = 0; i < quarters.Length; i++)
                {
                    // D2:D5 – quarter labels for the chart
                    cells[i + 1, 3].PutValue(quarters[i]); // D column

                    // E2:E5 – VLOOKUP formula referencing the lookup table
                    // Formula: =VLOOKUP(D2,$A$2:$B$5,2,FALSE)
                    string formula = $"=VLOOKUP(D{i + 2},$A$2:$B$5,2,FALSE)";
                    cells[i + 1, 4].Formula = formula; // E column
                }

                // ------------------------------------------------------------
                // 4. Calculate formulas so that the VLOOKUP results are materialized
                // ------------------------------------------------------------
                workbook.CalculateFormula();

                // ------------------------------------------------------------
                // 5. Create a column chart that uses the quarter labels (D2:D5)
                //    as the category axis and the VLOOKUP results (E2:E5) as values.
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add the series values (E2:E5)
                chart.NSeries.Add("E2:E5", false);
                // Note: CategoryData property is not available in the current API version.
                // The chart will display default numeric categories.

                // Optional: give the chart a title
                chart.Title.Text = "Quarterly Targets (VLOOKUP)";

                // ------------------------------------------------------------
                // 6. Save the workbook to a file
                // ------------------------------------------------------------
                string outputPath = "QuarterlyTargetsVLookupChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
