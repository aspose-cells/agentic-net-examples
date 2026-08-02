// Title: Aspose.Cells C# Example: VLOOKUP‑Driven Column Chart for Quarterly Targets
// Description: Demonstrates how to build a product lookup table, apply VLOOKUP formulas to fetch quarterly target values, create a column chart that reads those formula results, recalculate all formulas, and save the workbook as VLookupChartDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells VLOOKUP chart | C# Excel column chart from formula | dynamic chart data Aspose.Cells | calculate formulas Aspose.Cells | Excel VLOOKUP example C# | populate chart with lookup values
// Common Searches: how to use VLOOKUP with Aspose.Cells chart | C# create column chart from formula results | Aspose.Cells example VLOOKUP for chart data | generate Excel chart using lookup table in C# | Aspose.Cells calculate formulas before saving
// Developer Intent: Create an Excel workbook where a column chart displays quarterly targets retrieved through VLOOKUP formulas.
// Use Cases: Build a product‑wise quarterly target table and pull Q1 values into a chart via VLOOKUP. | Generate a chart that automatically updates when the source lookup table changes. | Ensure the chart shows numeric values by recalculating all formulas before saving.
// AI Prompts: Show how to modify the VLOOKUP formula to return Q2 or Q3 targets. | Add multiple series to the chart for Q1‑Q4 using VLOOKUP in Aspose.Cells. | Explain how to set chart title, axis labels, and legend programmatically after formula calculation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to build a product lookup table, apply VLOOKUP formulas to fetch quarterly target values, create a column chart that reads those formula results, recalculate all formulas, and save the workbook as VLookupChartDemo.xlsx using Aspose.Cells for .NET.
class VLookupChartDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Build a lookup table with quarterly targets
            // ------------------------------------------------------------
            // Header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Q1");
            cells["C1"].PutValue("Q2");
            cells["D1"].PutValue("Q3");
            cells["E1"].PutValue("Q4");

            // Sample data
            string[] products = { "ProdA", "ProdB", "ProdC" };
            int[,] targets = {
                { 100, 120, 130, 140 },
                { 200, 210, 220, 230 },
                { 300, 310, 320, 330 }
            };

            // Populate the lookup table (A2:E4)
            for (int i = 0; i < products.Length; i++)
            {
                cells[i + 1, 0].PutValue(products[i]);               // Column A: Product name
                for (int q = 0; q < 4; q++)
                    cells[i + 1, q + 1].PutValue(targets[i, q]);    // Columns B‑E: Quarterly targets
            }

            // ------------------------------------------------------------
            // 2. List of products to be displayed in the chart
            // ------------------------------------------------------------
            cells["G1"].PutValue("SelectedProduct");
            cells["G2"].PutValue("ProdA");
            cells["G3"].PutValue("ProdB");
            cells["G4"].PutValue("ProdC");

            // ------------------------------------------------------------
            // 3. Apply VLOOKUP formulas to fetch Q1 targets for each selected product
            //    Formula: =VLOOKUP(Gx,$A$2:$E$4,2,FALSE)
            // ------------------------------------------------------------
            for (int row = 2; row <= 4; row++)
            {
                string formula = $"=VLOOKUP(G{row},$A$2:$E$4,2,FALSE)";
                cells[$"H{row}"].SetFormula(formula, new FormulaParseOptions());
            }

            // ------------------------------------------------------------
            // 4. Create a column chart that uses the VLOOKUP results
            // ------------------------------------------------------------
            // Add a column chart positioned at rows 6‑20 and columns 0‑10
            int chartIdx = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIdx];

            // Series values come from H2:H4 (the VLOOKUP results)
            chart.NSeries.Add("H2:H4", true);
            // Category (X‑axis) labels come from G2:G4 (selected product names)
            chart.NSeries.CategoryData = "G2:G4";

            // ------------------------------------------------------------
            // 5. Calculate all formulas so that the chart reflects actual values
            // ------------------------------------------------------------
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 6. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("VLookupChartDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
