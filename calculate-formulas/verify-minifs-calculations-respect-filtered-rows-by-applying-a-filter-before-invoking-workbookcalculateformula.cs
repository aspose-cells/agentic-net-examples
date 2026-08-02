using System;
using Aspose.Cells;

class MinIfsFilteredDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate header
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");

        // Populate sample data (rows 2‑7)
        string[] categories = { "A", "B", "A", "B", "A", "B" };
        int[] values = { 10, 20, 5, 30, 15, 25 };
        for (int i = 0; i < categories.Length; i++)
        {
            cells[i + 1, 0].PutValue(categories[i]); // Column A
            cells[i + 1, 1].PutValue(values[i]);    // Column B
        }

        // Set MINIFS formula: minimum Value where Category = "A"
        cells["D1"].Formula = "=MINIFS(B2:B7, A2:A7, \"A\")";

        // Apply an AutoFilter to the data range and filter for Category = "A"
        worksheet.AutoFilter.Range = "A1:B7";          // Define the filter range (including header)
        worksheet.AutoFilter.Filter(0, "A");          // Column index 0 (Category) = "A"
        worksheet.AutoFilter.Refresh();               // Apply the filter (hides non‑matching rows)

        // Calculate all formulas after the filter has been applied
        workbook.CalculateFormula();

        // Output the MINIFS result – it should consider only the visible rows (Category = "A")
        Console.WriteLine("MINIFS result (visible rows only): " + cells["D1"].Value);

        // Optional: display which rows are hidden after filtering
        for (int row = 1; row <= 6; row++) // rows 2‑7 in the sheet (zero‑based index)
        {
            bool isHidden = cells.Rows[row].IsHidden;
            Console.WriteLine($"Row {row + 1} hidden: {isHidden}");
        }

        // Save the workbook to verify the filter and formula visually (optional)
        workbook.Save("MinIfsFilteredDemo.xlsx");
    }
}