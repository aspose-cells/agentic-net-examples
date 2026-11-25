// Title: Generate a column chart from VLOOKUP‑derived data range using Aspose.Cells for .NET
// Common Searches: how to add a VLOOKUP column to an Excel table with Aspose.Cells C# | Aspose.Cells C# create chart from formula results | C# generate column chart using VLOOKUP lookup values in Aspose.Cells | calculate formulas before creating chart with Aspose.Cells .NET | populate chart series from a lookup table using Aspose.Cells API
// Tags: VLOOKUP formula with Aspose.Cells | create Excel table and chart using Aspose.Cells | column chart from lookup results C# | calculate formulas before chart generation Aspose.Cells | Aspose.Cells chart series from dynamic range

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

// The example creates a workbook, defines a table of quarters and targets, adds a VLOOKUP column that retrieves targets for specified quarters, forces formula calculation, builds a column chart using the lookup results as series data, and saves the file as VlookupChart.xlsx.
class Program
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
            // 1. Populate source data: Quarter and Target
            // ------------------------------------------------------------
            cells["A1"].PutValue("Quarter");
            cells["B1"].PutValue("Target");

            string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
            double[] targets = { 120000, 150000, 130000, 160000 };

            for (int i = 0; i < quarters.Length; i++)
            {
                cells[i + 1, 0].PutValue(quarters[i]);   // Column A
                cells[i + 1, 1].PutValue(targets[i]);   // Column B
            }

            // Convert the source data into a ListObject (Excel Table)
            ListObjectCollection tables = worksheet.ListObjects;
            int tableIndex = tables.Add(0, 0, quarters.Length, 1, true);
            ListObject sourceTable = tables[tableIndex];
            sourceTable.DisplayName = "QuarterTargets";
            sourceTable.ShowHeaderRow = true;
            sourceTable.ShowTotals = false;

            // ------------------------------------------------------------
            // 2. Create a lookup area that uses VLOOKUP to retrieve targets
            // ------------------------------------------------------------
            cells["D1"].PutValue("LookupQuarter");
            cells["E1"].PutValue("RetrievedTarget");

            string[] lookupQuarters = { "Q1", "Q3", "Q5", "Q2" };
            for (int i = 0; i < lookupQuarters.Length; i++)
            {
                // Input quarter to look up
                cells[i + 1, 3].PutValue(lookupQuarters[i]); // Column D

                // VLOOKUP formula:
                // =VLOOKUP(D2,QuarterTargets!$A$2:$B$5,2,FALSE)
                string formula = $"=VLOOKUP(D{ i + 2 },QuarterTargets!$A$2:$B$5,2,FALSE)";
                cells[i + 1, 4].Formula = formula; // Column E
            }

            // Calculate all formulas so that the retrieved values are materialized
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 3. Add a chart that visualizes the VLOOKUP results
            // ------------------------------------------------------------
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 8, 0, 25, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.Title.Text = "Quarterly Targets via VLOOKUP";

            // Set the series values (E2:E5) and categories (D2:D5)
            chart.NSeries.Add("E2:E5", true);
            chart.NSeries.CategoryData = "D2:D5";

            // Optional chart formatting
            chart.Legend.Position = LegendPositionType.Bottom;

            // ------------------------------------------------------------
            // 4. Save the workbook
            // ------------------------------------------------------------
            string outputFile = "VlookupChart.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
