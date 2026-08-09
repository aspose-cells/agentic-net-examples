// Title: Aspose.Cells .NET Example: Build a VLOOKUP‑Driven Column Chart
// Description: Demonstrates how to create a workbook, set up a Quarter‑Target lookup table, apply VLOOKUP formulas, force calculation, and generate a column chart that visualizes the retrieved values using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# VLOOKUP | Excel VLOOKUP chart | Aspose.Cells column chart | formula calculation Aspose.Cells | dynamic chart data | Excel automation .NET | chart series from formulas | VLOOKUP example | Aspose.Cells workbook
// Common Searches: VLOOKUP formula in Aspose.Cells .NET | Create column chart from formula results Aspose.Cells | Calculate formulas before adding chart Aspose.Cells | Bind chart series to VLOOKUP cells Aspose.Cells | Aspose.Cells example for dynamic chart data
// Developer Intent: Generate an Excel file where a column chart shows quarterly targets fetched with VLOOKUP formulas via Aspose.Cells for .NET.
// Use Cases: Produce a quarterly target report that updates automatically when the lookup table changes. | Create a reusable Excel template that calculates values with VLOOKUP and visualizes them in a chart for financial presentations. | Automate batch generation of workbooks with different lookup ranges while ensuring formulas are evaluated before chart rendering.
// AI Prompts: Write C# code using Aspose.Cells to add a VLOOKUP formula range and bind the results to a column chart, then save the workbook. | Explain how to force formula calculation in Aspose.Cells before creating a chart so VLOOKUP results appear in the series. | Provide step‑by‑step instructions to set chart series and category data from cells containing VLOOKUP formulas in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace VLookupChartDemo
{
    // Demonstrates how to create a workbook, set up a Quarter‑Target lookup table, apply VLOOKUP formulas, force calculation, and generate a column chart that visualizes the retrieved values using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 1. Prepare the lookup table (Quarter -> Target)
                // ------------------------------------------------------------
                // Header
                cells["A1"].PutValue("Quarter");
                cells["B1"].PutValue("Target");

                // Data rows
                cells["A2"].PutValue("Q1");
                cells["B2"].PutValue(120000);
                cells["A3"].PutValue("Q2");
                cells["B3"].PutValue(150000);
                cells["A4"].PutValue("Q3");
                cells["B4"].PutValue(130000);
                cells["A5"].PutValue("Q4");
                cells["B5"].PutValue(170000);

                // ------------------------------------------------------------
                // 2. List of quarters for the chart (could be same as lookup keys)
                // ------------------------------------------------------------
                cells["D1"].PutValue("Quarter");
                cells["E1"].PutValue("Target (VLOOKUP)");

                cells["D2"].PutValue("Q1");
                cells["D3"].PutValue("Q2");
                cells["D4"].PutValue("Q3");
                cells["D5"].PutValue("Q4");

                // ------------------------------------------------------------
                // 3. Apply VLOOKUP formula to retrieve targets dynamically
                // ------------------------------------------------------------
                // Formula: =VLOOKUP(D2,$A$2:$B$5,2,FALSE)
                for (int row = 2; row <= 5; row++)
                {
                    string formula = $"=VLOOKUP(D{row},$A$2:$B$5,2,FALSE)";
                    cells[$"E{row}"].Formula = formula;
                }

                // ------------------------------------------------------------
                // 4. Calculate all formulas so that the VLOOKUP results are materialized
                // ------------------------------------------------------------
                workbook.CalculateFormula();

                // ------------------------------------------------------------
                // 5. Create a column chart that uses the quarters as categories
                //    and the VLOOKUP results as values
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series: values (E2:E5) and categories (D2:D5)
                chart.NSeries.Add("E2:E5", true);
                chart.NSeries.CategoryData = "D2:D5";

                // Optional: give the chart a title
                chart.Title.Text = "Quarterly Targets (VLOOKUP)";

                // ------------------------------------------------------------
                // 6. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "VLookupChart.xlsx";

                // Ensure the directory exists (in case a path is provided)
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine("Workbook with VLOOKUP-driven chart created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
