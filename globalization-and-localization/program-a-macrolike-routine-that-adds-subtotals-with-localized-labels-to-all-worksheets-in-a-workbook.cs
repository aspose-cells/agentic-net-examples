// Title: Create a macro‑like C# routine that adds a culture‑specific subtotal row with SUM formulas to every worksheet using Aspose.Cells
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets, inserts a subtotal row at the end of the data range, and uses the current UI culture to label the row. | Write a function that adds column‑wise SUM formulas, applies a bold light‑gray style to the new subtotal row, and saves the modified workbook.
// Common Searches: aspocells c# add subtotal row to each sheet with localized label | how to insert a culture‑aware subtotal row in an Excel workbook using Aspose.Cells | C# iterate all worksheets and calculate column totals with Aspose.Cells | apply custom style to subtotal row in Aspose.Cells workbook | generate sum formulas for each column automatically in Aspose.Cells .NET
// Tags: culture specific subtotal label Aspose.Cells | iterate worksheets Aspose.Cells C# | apply bold style to subtotal row Aspose.Cells | auto generate column sum formulas Aspose.Cells | culture aware Excel processing .NET

using System;
using System.Globalization;
using Aspose.Cells;

// The example loads an Excel workbook, iterates every worksheet, determines the used range, inserts a subtotal row with a label derived from the current UI culture, adds SUM formulas for each data column, styles the row with a bold light‑gray format, and saves the updated file.
class SubtotalAdder
{
    // Returns a localized label for "Subtotal" based on the current UI culture.
    private static string GetLocalizedSubtotalLabel()
    {
        // You can extend this method with a proper resource lookup.
        // For demonstration, we prepend the culture name.
        string culture = CultureInfo.CurrentCulture.TwoLetterISOLanguageName;
        return $"{culture.ToUpperInvariant()} Subtotal";
    }

    static void Main()
    {
        // Load the workbook (replace with your actual file path).
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range.
            int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
            int maxColumn = cells.MaxDataColumn; // zero‑based index of the last column with data

            // If the sheet is empty, skip it.
            if (maxRow < 0 || maxColumn < 0)
                continue;

            // Assume the first row (row 0) contains headers, data starts at row 1.
            int dataStartRow = 1;
            int totalRow = maxRow + 1; // Row where the subtotal will be placed.

            // Insert the localized label in the first column of the subtotal row.
            cells[totalRow, 0].PutValue(GetLocalizedSubtotalLabel());

            // Add SUM formulas for each data column (starting from column 1 to preserve the label column).
            for (int col = 1; col <= maxColumn; col++)
            {
                // Convert column index to Excel column letter (e.g., 0 -> A, 1 -> B).
                string colLetter = CellsHelper.ColumnIndexToName(col);

                // Build the SUM formula: =SUM(B2:B{maxRow+1})
                string formula = $"=SUM({colLetter}{dataStartRow + 1}:{colLetter}{maxRow + 1})";

                // Place the formula in the subtotal row.
                cells[totalRow, col].Formula = formula;
            }

            // Optional: Apply a style to the subtotal row for visual distinction.
            Style style = workbook.CreateStyle();
            style.ForegroundColor = System.Drawing.Color.LightGray;
            style.Pattern = BackgroundType.Solid;
            style.Font.IsBold = true;

            StyleFlag flag = new StyleFlag();
            flag.All = true;

            cells.CreateRange(totalRow, 0, 1, maxColumn + 1).ApplyStyle(style, flag);
        }

        // Save the modified workbook (replace with your desired output path).
        workbook.Save("output.xlsx");
    }
}
