// Title: Define a named range that references only visible rows using a filtered address in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that iterates through a column, skips hidden rows, builds a cell address list, and assigns it to a workbook named range. | Write a method that takes start and end cells, returns the addresses of all non‑hidden rows after an AutoFilter, and uses the result to set the RefersTo property of a workbook name.
// Common Searches: Aspose.Cells C# create named range that excludes hidden rows after applying AutoFilter | How to get addresses of visible cells in a filtered column using Aspose.Cells | C# build cell address list for visible rows in an Excel worksheet | Define a workbook name that points only to non‑hidden rows with Aspose.Cells .NET | Retrieve visible row addresses from AutoFilter and set a named range in Aspose.Cells
// Tags: Aspose.Cells named range from filtered rows | C# generate visible cell address list | AutoFilter skip hidden rows Aspose.Cells | Workbook Names.RefersTo set dynamic range | Excel .xlsx create named range programmatically

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace Example
{
    // The example creates a workbook, populates column A, hides specific rows, applies an AutoFilter, builds a comma‑separated address of the visible cells, adds a named range "VisibleValues" that points to that address, and saves the file as FilteredNamedRange.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with a header
                sheet.Cells["A1"].PutValue("Value");
                for (int i = 2; i <= 10; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);
                }

                // Hide specific rows (zero‑based index)
                sheet.Cells.Rows[3].IsHidden = true; // hides row 4
                sheet.Cells.Rows[6].IsHidden = true; // hides row 7

                // Apply AutoFilter to the range
                sheet.AutoFilter.Range = "A1:A10";
                sheet.AutoFilter.Refresh();

                // Build address of visible (non‑hidden) cells manually
                string filteredAddress = GetVisibleRange(sheet, "A1", "A10");

                // Add a named range that points to the visible cells
                int nameIdx = workbook.Worksheets.Names.Add("VisibleValues");
                // RefersTo must be a formula string, e.g., =Sheet1!A2,A3,...
                workbook.Worksheets.Names[nameIdx].RefersTo = $"={sheet.Name}!{filteredAddress}";

                // Save the workbook
                string outputPath = "FilteredNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Constructs a comma‑separated address of visible rows within the specified column range
        private static string GetVisibleRange(Worksheet sheet, string startCell, string endCell)
        {
            int startRow = CellReferenceHelper.CellNameToRow(startCell);
            int endRow = CellReferenceHelper.CellNameToRow(endCell);
            string columnName = CellReferenceHelper.ColumnIndexToName(CellReferenceHelper.CellNameToColumn(startCell));

            var sb = new StringBuilder();
            bool first = true;
            for (int row = startRow; row <= endRow; row++)
            {
                // Skip hidden rows
                if (sheet.Cells.Rows[row].IsHidden)
                    continue;

                if (!first)
                    sb.Append(",");
                sb.Append($"{columnName}{row + 1}");
                first = false;
            }
            return sb.ToString();
        }
    }

    // Helper for converting between cell references and indices
    internal static class CellReferenceHelper
    {
        public static int CellNameToRow(string cellName)
        {
            string rowPart = Regex.Match(cellName, @"\d+").Value;
            return int.Parse(rowPart) - 1; // zero‑based
        }

        public static int CellNameToColumn(string cellName)
        {
            string colPart = Regex.Match(cellName, @"[A-Za-z]+").Value.ToUpper();
            int sum = 0;
            foreach (char c in colPart)
            {
                sum = sum * 26 + (c - 'A' + 1);
            }
            return sum - 1; // zero‑based
        }

        public static string ColumnIndexToName(int index)
        {
            index++; // convert to 1‑based
            string name = "";
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                name = (char)('A' + rem) + name;
                index = (index - 1) / 26;
            }
            return name;
        }
    }
}
