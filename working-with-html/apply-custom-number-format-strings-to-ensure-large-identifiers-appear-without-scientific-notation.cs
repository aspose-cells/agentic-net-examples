using Aspose.Cells;

// Author: Aspose.Cells .NET example – applying custom number format to prevent scientific notation
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual load logic as per the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Apply custom format to a range containing large identifiers
        for (int row = 0; row < 10; row++)               // example: first 10 rows in column A
        {
            Cell cell = sheet.Cells[row, 0];
            cell.PutValue(1234567890123456L);            // a 16‑digit identifier

            // Retrieve the cell's current style
            Style style = cell.GetStyle();

            // Set a custom number format that forces full integer display (no scientific notation)
            style.SetCustom("0", false);                 // "0" = integer without grouping or exponent

            // Prepare a StyleFlag to indicate that only the number format should be applied
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the cell
            cell.SetStyle(style, flag);
        }

        // Optionally, enforce a global significant‑digits setting (e.g., 15 digits) to avoid rounding
        CellsHelper.SignificantDigitsType = SignificantDigitsType.Digits15;

        // Save the workbook (replace with your actual save logic as per the provided save rule)
        workbook.Save("output.xlsx");
    }
}