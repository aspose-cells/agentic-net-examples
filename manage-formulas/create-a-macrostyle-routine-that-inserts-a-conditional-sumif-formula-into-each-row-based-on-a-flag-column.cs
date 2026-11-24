// Title: Insert a row‑by‑row conditional SUMIF formula using an Aspose.Cells macro‑style routine in C#
// Common Searches: Aspose.Cells C# insert SUMIF formula for each row based on a flag column | how to generate conditional SUMIF formulas programmatically with Aspose.Cells | macro style routine to add SUMIF per row in Excel using Aspose.Cells .NET | C# Aspose.Cells set formula in column D that sums amounts where flag equals current row | create dynamic SUMIF range with Aspose.Cells workbook in C#
// Tags: Aspose.Cells insert SUMIF formula | C# generate conditional Excel formulas | macro‑style worksheet automation Aspose.Cells | row‑wise SUMIF calculation .NET | dynamic range construction with CellsHelper

using System;
using Aspose.Cells;

namespace AsposeCellsMacroStyle
{
    // The example creates a workbook, fills a flag column and an amount column, then uses a macro‑style method to loop through each data row and assign a SUMIF formula that sums amounts matching the row's flag, placing the formula in a target column and saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: Flag column (A), Values column (B)
            // Header row
            cells["A1"].PutValue("Flag");
            cells["B1"].PutValue("Amount");
            cells["D1"].PutValue("Conditional SumIf");

            // Populate sample rows (for demonstration)
            string[] flags = { "Yes", "No", "Yes", "No", "Yes" };
            double[] amounts = { 100, 200, 150, 250, 300 };

            for (int i = 0; i < flags.Length; i++)
            {
                int row = i + 2; // Data starts at row 2
                cells[row, 0].PutValue(flags[i]);   // Column A
                cells[row, 1].PutValue(amounts[i]); // Column B
            }

            // Insert conditional SUMIF formulas into column D
            InsertConditionalSumIf(sheet, startDataRow: 2, endDataRow: flags.Length + 1,
                                   flagColumnIndex: 0, amountColumnIndex: 1, targetColumnIndex: 3);

            // Save the workbook
            workbook.Save("ConditionalSumIf.xlsx");
        }

        /// <param name="sheet">Worksheet to operate on.</param>
        /// <param name="startDataRow">First data row index (1‑based).</param>
        /// <param name="endDataRow">Last data row index (1‑based).</param>
        /// <param name="flagColumnIndex">Zero‑based index of the flag column (e.g., 0 for column A).</param>
        /// <param name="amountColumnIndex">Zero‑based index of the amount column (e.g., 1 for column B).</param>
        /// <param name="targetColumnIndex">Zero‑based index where the formula will be placed (e.g., 3 for column D).</param>
        static void InsertConditionalSumIf(Worksheet sheet, int startDataRow, int endDataRow,
                                           int flagColumnIndex, int amountColumnIndex, int targetColumnIndex)
        {
            Cells cells = sheet.Cells;

            // Build absolute range strings for the flag and amount columns
            string flagRange = $"${CellIndexToName(flagColumnIndex)}${startDataRow}:${CellIndexToName(flagColumnIndex)}${endDataRow}";
            string amountRange = $"${CellIndexToName(amountColumnIndex)}${startDataRow}:${CellIndexToName(amountColumnIndex)}${endDataRow}";

            // Loop through each row and set the SUMIF formula
            for (int row = startDataRow; row <= endDataRow; row++)
            {
                // Criteria is the flag cell of the current row (relative reference)
                string criteriaCell = $"{CellIndexToName(flagColumnIndex)}{row}";

                // Construct the SUMIF formula
                string formula = $"=SUMIF({flagRange}, {criteriaCell}, {amountRange})";

                // Place the formula in the target column of the current row
                cells[row, targetColumnIndex].Formula = formula;
            }
        }

        /// <summary>
        /// Converts a zero‑based column index to its Excel column name (e.g., 0 → "A").
        /// </summary>
        static string CellIndexToName(int columnIndex)
        {
            return CellsHelper.ColumnIndexToName(columnIndex);
        }
    }
}
