// Title: Compare cell values of two HTML workbooks before and after redundant space deletion using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells compare two HTML workbooks after removing redundant spaces | how to detect cell differences between original and cleaned HTML files with Aspose.Cells | compare HTML worksheets cell by cell using Aspose.Cells in a C# console app | find discrepancies in HTML workbooks generated before and after space cleanup | Aspose.Cells C# example for comparing HTML output before/after space optimization
// Tags: compare html workbooks Aspose.Cells C# | detect cell value differences Aspose.Cells HTML | load html file into Aspose.Cells workbook | iterate worksheets and cells Aspose.Cells | report discrepancies in HTML worksheet data C# | handle missing html files Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the presence of two HTML files, loads each into an Aspose.Cells Workbook, determines the overlapping worksheet count, and iterates through the maximum used rows and columns. It compares each cell's string value, outputs any differences to the console, and includes robust error handling for missing files or unexpected exceptions.
class HtmlComparison
{
    static void Main()
    {
        try
        {
            const string beforePath = "Before.html";
            const string afterPath = "After.html";

            // Verify that the required HTML files exist before attempting to load them
            if (!File.Exists(beforePath))
            {
                Console.WriteLine($"File not found: {beforePath}");
                return;
            }

            if (!File.Exists(afterPath))
            {
                Console.WriteLine($"File not found: {afterPath}");
                return;
            }

            // Load the HTML files into workbooks
            Workbook wbBefore = new Workbook(beforePath);
            Workbook wbAfter = new Workbook(afterPath);

            // Use the smaller worksheet count to avoid out‑of‑range errors
            int sheetCount = Math.Min(wbBefore.Worksheets.Count, wbAfter.Worksheets.Count);

            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet wsBefore = wbBefore.Worksheets[i];
                Worksheet wsAfter = wbAfter.Worksheets[i];

                // Determine the maximum used row and column across both worksheets
                int maxRow = Math.Max(wsBefore.Cells.MaxDataRow, wsAfter.Cells.MaxDataRow);
                int maxCol = Math.Max(wsBefore.Cells.MaxDataColumn, wsAfter.Cells.MaxDataColumn);

                // Iterate through each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve cell values; treat null cells as empty strings
                        string valBefore = wsBefore.Cells[row, col]?.StringValue ?? string.Empty;
                        string valAfter = wsAfter.Cells[row, col]?.StringValue ?? string.Empty;

                        // Report any differences
                        if (valBefore != valAfter)
                        {
                            Console.WriteLine($"Difference in sheet '{wsBefore.Name}' cell [{row},{col}]:");
                            Console.WriteLine($"  Before: \"{valBefore}\"");
                            Console.WriteLine($"  After : \"{valAfter}\"");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
