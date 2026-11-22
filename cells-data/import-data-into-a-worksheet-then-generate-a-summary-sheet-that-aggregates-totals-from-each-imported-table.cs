// Title: C# example: import multiple DataTables into separate worksheets and create a summary sheet with SUMIF aggregation using Aspose.Cells
// Common Searches: how to import a DataTable into a new worksheet with Aspose.Cells in C# | Aspose.Cells C# code to generate a summary worksheet that sums values from several sheets | using SUMIF formulas in Aspose.Cells to total quantities across multiple Excel tabs | calculate average price across Q1 and Q2 sales sheets with Aspose.Cells C# example | C# Aspose.Cells tutorial for creating an aggregated report workbook with several worksheets
// Tags: import datatable to worksheet Aspose.Cells | create summary worksheet with SUMIF Aspose.Cells | aggregate totals across multiple worksheets Aspose.Cells | average price calculation using worksheet formulas Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSummaryExample
{
    // The sample creates a new Workbook, imports two DataTables into worksheets named "SalesQ1" and "SalesQ2" using a helper method, then adds a "Summary" worksheet that lists products and applies SUMIF formulas to total quantities and compute average prices across the sales sheets. After recalculating formulas, the workbook is saved as AggregatedReport.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Import first data table into worksheet "SalesQ1"
                // -------------------------------------------------
                Worksheet salesQ1Sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                salesQ1Sheet.Name = "SalesQ1";

                // Sample data for SalesQ1
                DataTable dtSalesQ1 = new DataTable();
                dtSalesQ1.Columns.Add("Product", typeof(string));
                dtSalesQ1.Columns.Add("Quantity", typeof(int));
                dtSalesQ1.Columns.Add("Price", typeof(double));

                dtSalesQ1.Rows.Add("Apple", 120, 0.5);
                dtSalesQ1.Rows.Add("Banana", 80, 0.3);
                dtSalesQ1.Rows.Add("Orange", 150, 0.4);

                // Import DataTable starting at cell A1 (row 0, column 0)
                WriteDataTableToSheet(salesQ1Sheet, dtSalesQ1, 0, 0, includeColumnNames: true);

                // -------------------------------------------------
                // 2. Import second data table into worksheet "SalesQ2"
                // -------------------------------------------------
                Worksheet salesQ2Sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                salesQ2Sheet.Name = "SalesQ2";

                // Sample data for SalesQ2
                DataTable dtSalesQ2 = new DataTable();
                dtSalesQ2.Columns.Add("Product", typeof(string));
                dtSalesQ2.Columns.Add("Quantity", typeof(int));
                dtSalesQ2.Columns.Add("Price", typeof(double));

                dtSalesQ2.Rows.Add("Apple", 130, 0.55);
                dtSalesQ2.Rows.Add("Banana", 90, 0.35);
                dtSalesQ2.Rows.Add("Orange", 160, 0.45);

                // Import DataTable starting at cell A1 (row 0, column 0)
                WriteDataTableToSheet(salesQ2Sheet, dtSalesQ2, 0, 0, includeColumnNames: true);

                // -------------------------------------------------
                // 3. Create summary worksheet that aggregates totals
                // -------------------------------------------------
                Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
                summarySheet.Name = "Summary";

                // Header row
                summarySheet.Cells["A1"].PutValue("Product");
                summarySheet.Cells["B1"].PutValue("Total Quantity");
                summarySheet.Cells["C1"].PutValue("Average Price");

                // List of products (assumed same across tables)
                string[] products = { "Apple", "Banana", "Orange" };
                for (int i = 0; i < products.Length; i++)
                {
                    int row = i + 2; // Data starts from row 2 (1‑based index)
                    summarySheet.Cells[row, 0].PutValue(products[i]);

                    // Formula to sum Quantity from both tables for the current product
                    string qtyFormula = $"=SUMIF(SalesQ1!A:A,\"{products[i]}\",SalesQ1!B:B)+SUMIF(SalesQ2!A:A,\"{products[i]}\",SalesQ2!B:B)";
                    summarySheet.Cells[row, 1].Formula = qtyFormula;

                    // Formula to calculate average Price across both tables for the current product
                    string priceFormula = $"=(SUMIF(SalesQ1!A:A,\"{products[i]}\",SalesQ1!C:C)+SUMIF(SalesQ2!A:A,\"{products[i]}\",SalesQ2!C:C))/2";
                    summarySheet.Cells[row, 2].Formula = priceFormula;
                }

                // Recalculate all formulas so that values are stored in the file
                workbook.CalculateFormula();

                // Save the workbook (lifecycle rule: save)
                string outputPath = "AggregatedReport.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void WriteDataTableToSheet(Worksheet sheet, DataTable table, int startRow, int startColumn, bool includeColumnNames)
        {
            int rowIdx = startRow;
            int colIdx = startColumn;

            // Write column headers if required
            if (includeColumnNames)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    sheet.Cells[rowIdx, colIdx + c].PutValue(table.Columns[c].ColumnName);
                }
                rowIdx++; // Move to next row for data
            }

            // Write data rows
            foreach (DataRow dr in table.Rows)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    sheet.Cells[rowIdx, colIdx + c].PutValue(dr[c]);
                }
                rowIdx++;
            }
        }
    }
}
