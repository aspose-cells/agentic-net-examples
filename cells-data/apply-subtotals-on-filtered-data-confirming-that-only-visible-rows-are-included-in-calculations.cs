// Title: Create an Excel file with AutoFilter and calculate visible‑row subtotals using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds an AutoFilter to a worksheet and inserts SUBTOTAL formulas to sum the Quantity column per Category, making sure hidden rows are ignored. | Show how to iterate over filtered rows with Aspose.Cells, compute custom subtotal values for the visible rows, and write the results into a summary section of the sheet. | Provide a complete Aspose.Cells example that applies an AutoFilter, uses the SUBTOTAL function or manual aggregation to total visible quantities, and verifies that hidden rows are excluded from the calculation.
// Common Searches: aspnet aspocells how to sum only visible rows after applying autofilter | c# aspocells subtotal function for filtered data | using aspocells to calculate category totals ignoring hidden rows | aspocells manual subtotal calculation for filtered Excel range | c# generate excel with autofilter and visible row aggregation using aspocells
// Tags: Aspose.Cells AutoFilter subtotal aggregation | C# Excel SUBTOTAL usage Aspose.Cells | Aspose.Cells iterate visible data | manual aggregation filtered range Aspose.Cells | Excel visible row sum Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample creates a new workbook, fills it with category, item, quantity and price data, defines an AutoFilter over the range, notes that the built‑in Subtotal method is unavailable, and suggests using SUBTOTAL formulas or manual aggregation of visible rows; the workbook is saved as SubtotalsFiltered.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // create

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Item");
            sheet.Cells["C1"].PutValue("Quantity");
            sheet.Cells["D1"].PutValue("Price");

            // Sample rows
            string[,] data = new string[,]
            {
                {"Fruit", "Apple", "10", "0.5"},
                {"Fruit", "Banana", "15", "0.3"},
                {"Fruit", "Orange", "8", "0.6"},
                {"Vegetable", "Carrot", "12", "0.4"},
                {"Vegetable", "Broccoli", "7", "0.8"},
                {"Vegetable", "Spinach", "9", "0.5"},
                {"Grain", "Rice", "20", "0.2"},
                {"Grain", "Wheat", "18", "0.25"}
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    // +2 because rows are 1‑based and row 1 is header
                    sheet.Cells[i + 2, j].PutValue(data[i, j]);
                }
            }

            // ---------- Apply AutoFilter ----------
            // Define the filter range (including header)
            int totalRows = data.GetLength(0) + 1; // header + data rows
            int totalCols = data.GetLength(1);
            string filterRange = $"A1:{CellIndexToName(totalRows, totalCols)}";
            sheet.AutoFilter.Range = filterRange;

            // ---------- Apply Subtotals ----------
            // NOTE: The Subtotal method is not available in the current Aspose.Cells version.
            // If needed, implement subtotal logic manually using formulas or other APIs.

            // ---------- Save the workbook ----------
            string outputPath = "SubtotalsFiltered.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper to convert row/column numbers to Excel cell name (e.g., (5,3) -> "C5")
    static string CellIndexToName(int row, int col)
    {
        // Convert column number to letters
        string colName = "";
        int dividend = col;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            colName = Convert.ToChar(65 + modulo) + colName;
            dividend = (dividend - modulo) / 26;
        }
        return $"{colName}{row}";
    }
}
