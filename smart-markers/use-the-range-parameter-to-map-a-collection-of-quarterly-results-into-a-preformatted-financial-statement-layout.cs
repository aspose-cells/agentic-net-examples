// Title: Use Aspose.Cells Range to bulk‑insert quarterly revenue data into a pre‑formatted financial statement worksheet (C#)
// AI Prompts: Write C# code that creates a workbook, adds a header row, and uses Cells.CreateRange to assign a two‑dimensional array of quarterly revenue values to the worksheet in a single operation. | Generate a method that accepts a List<QuarterlyResult>, fills the year column, inserts the quarterly numbers via a Range.Value assignment, adds a SUM formula for each row to compute the total, and saves the workbook. | Extend the example to include a cumulative total column that sums the yearly totals across all years using Aspose.Cells formulas.
// Common Searches: how to bulk load a list of objects into Excel using Aspose.Cells CreateRange C# | populate a financial statement template with quarterly data using Aspose.Cells | add row‑wise SUM formulas after inserting data with Aspose.Cells Range | map collection to Excel cells with a single Range.Value assignment in C# | calculate totals per row in Aspose.Cells after bulk data insertion
// Tags: Aspose.Cells CreateRange bulk assignment | C# map quarterly results to Excel worksheet | row total calculation with SUM formula Aspose.Cells | automate financial statement layout Excel | populate Excel header row using Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace FinancialStatementMapping
{
    // Simple model for quarterly results of a given year
    // Demonstrates creating a workbook, building a header row, using Aspose.Cells.Range to bulk‑insert quarterly revenue values, writing the year column, adding per‑row SUM formulas for totals, recalculating formulas, and saving the file as FinancialStatement.xlsx.
    public class QuarterlyResult
    {
        public int Year { get; set; }
        public double[] Quarters { get; set; } = new double[4]; // Q1..Q4
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Sample data: quarterly revenues for three years
                List<QuarterlyResult> data = new List<QuarterlyResult>
                {
                    new QuarterlyResult { Year = 2021, Quarters = new double[] { 12000, 15000, 13000, 16000 } },
                    new QuarterlyResult { Year = 2022, Quarters = new double[] { 14000, 15500, 14500, 17000 } },
                    new QuarterlyResult { Year = 2023, Quarters = new double[] { 15000, 16500, 15500, 18000 } }
                };

                // -------------------------------------------------
                // 1. Create a new workbook (lifecycle rule)
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // 2. Build the pre‑formatted layout
                // -------------------------------------------------
                // Header row
                cells["A1"].PutValue("Year");
                cells["B1"].PutValue("Q1");
                cells["C1"].PutValue("Q2");
                cells["D1"].PutValue("Q3");
                cells["E1"].PutValue("Q4");
                cells["F1"].PutValue("Total");

                // Apply a simple style to the header (optional)
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;
                cells.CreateRange("A1", "F1").SetStyle(headerStyle);

                // -------------------------------------------------
                // 3. Map the collection into the worksheet using Range
                // -------------------------------------------------
                // Prepare a 2‑dimensional object array for the quarterly values
                object[][] quarterValues = new object[data.Count][];
                for (int i = 0; i < data.Count; i++)
                {
                    quarterValues[i] = new object[4];
                    for (int q = 0; q < 4; q++)
                    {
                        quarterValues[i][q] = data[i].Quarters[q];
                    }
                }

                // Destination range for the quarterly numbers (starts at B2)
                // firstRow = 1 (zero‑based, i.e., row 2), firstColumn = 1 (column B)
                // totalRows = data.Count, totalColumns = 4
                Aspose.Cells.Range destQuarterRange = cells.CreateRange(1, 1, data.Count, 4);
                // Set the whole block in one operation
                destQuarterRange.Value = quarterValues;

                // Write the year column separately (simple loop)
                for (int i = 0; i < data.Count; i++)
                {
                    cells[i + 1, 0].PutValue(data[i].Year); // Column A
                }

                // -------------------------------------------------
                // 4. Calculate totals per row using a formula
                // -------------------------------------------------
                // The total column starts at column F (index 5)
                for (int i = 0; i < data.Count; i++)
                {
                    // Formula: =SUM(Bx:Ex) where x = row index + 2 (because rows are 1‑based in Excel)
                    int excelRow = i + 2;
                    cells[i + 1, 5].Formula = $"=SUM(B{excelRow}:E{excelRow})";
                }

                // Recalculate to materialize the totals
                workbook.CalculateFormula();

                // -------------------------------------------------
                // 5. Save the workbook (lifecycle rule)
                // -------------------------------------------------
                workbook.Save("FinancialStatement.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
