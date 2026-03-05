using System;
using Aspose.Cells;
using System.Drawing;

class FindCellDemo
{
    static void Main()
    {
        // Load an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Example 1: Find a cell that exactly matches "Orange"
        // -------------------------------------------------
        Cell previousCell = null; // start search from the beginning
        Cell exactMatch = worksheet.Cells.Find("Orange", previousCell);

        if (exactMatch != null)
        {
            Console.WriteLine($"Exact match found at {exactMatch.Name}");
            // Highlight the found cell
            Style highlight = workbook.CreateStyle();
            highlight.Font.Color = Color.Red;
            highlight.Font.IsBold = true;
            exactMatch.SetStyle(highlight);
        }
        else
        {
            Console.WriteLine("Exact match for \"Orange\" not found.");
        }

        // -------------------------------------------------
        // Example 2: Find a cell that contains the substring "app"
        // (case‑insensitive search in cell values)
        // -------------------------------------------------
        FindOptions options = new FindOptions
        {
            LookInType = LookInType.Values,   // search in cell values
            LookAtType = LookAtType.Contains   // value contains the search key
        };

        Cell containsMatch = worksheet.Cells.Find("app", null, options);

        if (containsMatch != null)
        {
            Console.WriteLine($"Contains match found at {containsMatch.Name} with value \"{containsMatch.StringValue}\"");
        }
        else
        {
            Console.WriteLine("No cell containing \"app\" was found.");
        }

        // -------------------------------------------------
        // Save the workbook (with any applied styling)
        // -------------------------------------------------
        workbook.Save("output.xlsx");
    }
}