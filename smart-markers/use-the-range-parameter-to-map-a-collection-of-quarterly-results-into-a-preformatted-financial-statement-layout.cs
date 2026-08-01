// Title: C# – Map a List of quarterly results to a formatted financial statement using Aspose.Cells Range
// Description: Creates a workbook, builds a header, fills a hidden source Range with a List<QuarterlyResult>, copies it to the visible layout, adds a grand‑total row with SUM formulas, and saves the file as FinancialStatement.xlsx.
// Keywords: Aspose.Cells C# | Range.CopyValue | map collection to Excel | financial statement layout | quarterly results POCO | temporary source range | Excel export .NET | SUM formula in Aspose.Cells
// Common Searches: Aspose.Cells copy range to another range C# | populate Excel table from List<QuarterlyResult> | add total row with SUM formulas using Aspose.Cells | use hidden source range for Excel export | map POCO collection to pre‑formatted worksheet
// Developer Intent: Transfer quarterly data from a POCO collection into a pre‑designed worksheet by leveraging a source Range and copying its values to the target layout.
// Use Cases: Generate a quarterly financial report by loading data into a hidden range and mapping it to a printable table. | Separate raw data storage from presentation layer to keep the layout clean while still using the same workbook. | Append a calculated grand‑total row after copying data, ensuring formulas update automatically.
// AI Prompts: Write C# code that creates a source Range from a List<QuarterlyResult> and copies it to a target Range with Aspose.Cells, preserving number formats. | Show how to add a grand‑total row with SUM formulas after copying data via Range in an Aspose.Cells workbook. | Explain the workflow for using a hidden source range to prepare data before mapping it to a formatted financial statement layout.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace FinancialStatementMapping
{
    // Simple POCO to hold quarterly results for a given year
    // Creates a workbook, builds a header, fills a hidden source Range with a List<QuarterlyResult>, copies it to the visible layout, adds a grand‑total row with SUM formulas, and saves the file as FinancialStatement.xlsx.
    public class QuarterlyResult
    {
        public int Year { get; set; }
        public double Q1 { get; set; }
        public double Q2 { get; set; }
        public double Q3 { get; set; }
        public double Q4 { get; set; }

        public double Total => Q1 + Q2 + Q3 + Q4;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Create a new workbook (lifecycle rule) ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- 2. Build a pre‑formatted financial statement layout ----------
                // Header row
                string[] headers = { "Year", "Q1", "Q2", "Q3", "Q4", "Total" };
                for (int col = 0; col < headers.Length; col++)
                {
                    cells[0, col].PutValue(headers[col]);
                    // Simple formatting: bold header
                    Style style = cells[0, col].GetStyle();
                    style.Font.IsBold = true;
                    cells[0, col].SetStyle(style);
                }

                // Set column widths for better readability
                for (int col = 0; col < headers.Length; col++)
                    cells.SetColumnWidth(col, 12);

                // ---------- 3. Prepare a collection of quarterly results ----------
                List<QuarterlyResult> results = new List<QuarterlyResult>
                {
                    new QuarterlyResult { Year = 2021, Q1 = 12000, Q2 = 15000, Q3 = 13000, Q4 = 16000 },
                    new QuarterlyResult { Year = 2022, Q1 = 14000, Q2 = 15500, Q3 = 13500, Q4 = 17000 },
                    new QuarterlyResult { Year = 2023, Q1 = 15000, Q2 = 16000, Q3 = 14000, Q4 = 18000 }
                };

                // ---------- 4. Populate a temporary source range with the collection ----------
                // We'll place the raw data starting at cell Z1 (far away from the layout)
                int sourceStartRow = 0;   // zero‑based index (row 1)
                int sourceStartCol = 25;  // column Z (0‑based)
                int rows = results.Count;
                int cols = 6; // Year + 4 quarters + Total

                // Create the source range
                Aspose.Cells.Range sourceRange = cells.CreateRange(sourceStartRow, sourceStartCol, rows, cols);

                // Fill the source range with data from the collection
                for (int i = 0; i < rows; i++)
                {
                    QuarterlyResult r = results[i];
                    sourceRange[i, 0].PutValue(r.Year);
                    sourceRange[i, 1].PutValue(r.Q1);
                    sourceRange[i, 2].PutValue(r.Q2);
                    sourceRange[i, 3].PutValue(r.Q3);
                    sourceRange[i, 4].PutValue(r.Q4);
                    sourceRange[i, 5].PutValue(r.Total);
                }

                // ---------- 5. Map the source range into the pre‑formatted layout ----------
                // Target range starts at row 2 (index 1) under the header, column A (index 0)
                int targetStartRow = 1;
                int targetStartCol = 0;
                Aspose.Cells.Range targetRange = cells.CreateRange(targetStartRow, targetStartCol, rows, cols);

                // Copy values (including number formats) from source to target
                targetRange.CopyValue(sourceRange);

                // ---------- 6. Optional: Add a simple total row using a formula ----------
                int totalRowIndex = targetStartRow + rows; // row after the last data row
                cells[totalRowIndex, 0].PutValue("Grand Total");
                // Apply bold style to the label
                Style totalLabelStyle = cells[totalRowIndex, 0].GetStyle();
                totalLabelStyle.Font.IsBold = true;
                cells[totalRowIndex, 0].SetStyle(totalLabelStyle);

                // Formula to sum each quarter column
                for (int col = 1; col <= 5; col++) // Q1..Total columns
                {
                    string colLetter = CellsHelper.ColumnIndexToName(col);
                    string formula = $"=SUM({colLetter}2:{colLetter}{rows + 1})";
                    cells[totalRowIndex, col].Formula = formula;
                }

                // ---------- 7. Save the workbook (lifecycle rule) ----------
                workbook.Save("FinancialStatement.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
