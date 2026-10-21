// Title: How to replace whole‑column references with specific range formulas in Excel using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace A:A column reference with A1:A1000 range | programmatically limit Excel formula ranges to used rows with Aspose.Cells | C# code to convert whole column formulas to bounded ranges in a workbook | optimize Excel calculation speed by changing column references via Aspose.Cells API | detect and modify column‑wide formulas in .xlsx using Aspose.Cells .NET
// Tags: replace whole column references Aspose.Cells | convert column formulas to bounded ranges C# | limit formula range to last used row .NET | optimize Excel calculation performance Aspose.Cells | regex based formula update C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, determines the last used row for each worksheet, and uses a regular expression to find whole‑column references (e.g., A:A) in formula cells. Each match is replaced with a bounded range (e.g., A1:A1000), the cell formula is updated only when changed, and the workbook is saved. The code also handles missing files, password‑protected workbooks, and other exceptions.
class ReplaceColumnFormulas
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook; if the file is password‑protected, an exception will be thrown
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last used row (1‑based index)
                int lastRow = sheet.Cells.MaxDataRow + 1;
                if (lastRow < 1) lastRow = 1; // Ensure at least row 1 exists

                // Regex to match whole‑column references like A:A, AB:AB, etc.
                Regex wholeColumnRegex = new Regex(@"([A-Z]+):\1", RegexOptions.IgnoreCase);

                // Iterate over all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula) continue;

                    string originalFormula = cell.Formula;

                    // Replace whole‑column references with a specific range (e.g., A1:A1000)
                    string newFormula = wholeColumnRegex.Replace(originalFormula, match =>
                    {
                        string column = match.Groups[1].Value;
                        return $"{column}1:{column}{lastRow}";
                    });

                    // Update the cell only if the formula changed
                    if (!originalFormula.Equals(newFormula, StringComparison.Ordinal))
                    {
                        cell.Formula = newFormula;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected workbook or other Cells-specific errors
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password to load it.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
