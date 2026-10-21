// Title: Generate a Summary Sheet that Lists Formula Counts by Function for Each Worksheet Using Aspose.Cells in C#
// Common Searches: how to create a summary worksheet that shows how many times each Excel function is used per sheet with Aspose.Cells C# | Aspose.Cells C# count formulas by function across multiple worksheets and write results to a new sheet | extract function name from Excel formula and aggregate counts using Aspose.Cells in .NET | C# code to add a summary tab with function usage statistics in an existing workbook using Aspose.Cells | automatically generate a formula usage report in Excel with Aspose.Cells library
// Tags: formula function count aggregation Aspose.Cells | parse Excel formula to obtain function name C# | add summary worksheet with aggregated data Aspose.Cells | scan used range for formulas Aspose.Cells | auto-fit columns after writing summary Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (creating a sample if none exists), iterates through each worksheet except any named "Summary", scans the used cell range to identify formula cells, parses each formula to extract its function name, tallies the occurrences per function per sheet, creates a new "Summary" worksheet with columns Worksheet, Function, and Count, writes the aggregated data, auto‑fits the columns, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Ensure the input file exists; if not, create a simple workbook for demonstration
            if (!File.Exists(inputPath))
            {
                workbook = CreateSampleWorkbook();
                workbook.Save(inputPath);
            }
            else
            {
                // Load the existing workbook (no password assumed)
                workbook = new Workbook(inputPath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Dictionary to store counts: Worksheet name -> (Function name -> Count)
        var sheetFunctionCounts = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        // Iterate through each worksheet (excluding any existing Summary sheet)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                continue;

            var functionCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Scan all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsFormula)
                    {
                        // Extract the function name from the formula string
                        string functionName = ExtractFunctionName(cell.Formula);
                        if (!string.IsNullOrEmpty(functionName))
                        {
                            if (functionCounts.ContainsKey(functionName))
                                functionCounts[functionName]++;
                            else
                                functionCounts[functionName] = 1;
                        }
                    }
                }
            }

            sheetFunctionCounts[sheet.Name] = functionCounts;
        }

        // Remove existing Summary sheet if present, then add a new one
        Worksheet existingSummary = workbook.Worksheets["Summary"];
        if (existingSummary != null)
            workbook.Worksheets.RemoveAt(existingSummary.Index);
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");

        // Write header row
        summarySheet.Cells[0, 0].PutValue("Worksheet");
        summarySheet.Cells[0, 1].PutValue("Function");
        summarySheet.Cells[0, 2].PutValue("Count");

        int currentRow = 1;

        // Populate the summary data
        foreach (var sheetEntry in sheetFunctionCounts)
        {
            string sheetName = sheetEntry.Key;
            foreach (var funcEntry in sheetEntry.Value)
            {
                summarySheet.Cells[currentRow, 0].PutValue(sheetName);
                summarySheet.Cells[currentRow, 1].PutValue(funcEntry.Key);
                summarySheet.Cells[currentRow, 2].PutValue(funcEntry.Value);
                currentRow++;
            }
        }

        // Adjust column widths for readability
        summarySheet.AutoFitColumns();

        try
        {
            // Save the workbook with the new Summary sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Summary sheet created successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    // Creates a simple workbook with a few formulas for demonstration purposes
    private static Workbook CreateSampleWorkbook()
    {
        var wb = new Workbook();
        var ws = wb.Worksheets[0];
        ws.Name = "Data";

        ws.Cells["A1"].PutValue(10);
        ws.Cells["A2"].PutValue(20);
        ws.Cells["A3"].Formula = "=SUM(A1:A2)";
        ws.Cells["A4"].Formula = "=AVERAGE(A1:A2)";
        ws.Cells["A5"].Formula = "=MAX(A1:A2)";

        return wb;
    }

    // Helper to extract the function name from a formula string (e.g., "=SUM(A1:A10)" -> "SUM")
    private static string ExtractFunctionName(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return null;

        // Remove leading '=' if present
        if (formula.StartsWith("="))
            formula = formula.Substring(1);

        int openParenIndex = formula.IndexOf('(');
        if (openParenIndex <= 0)
            return null;

        // Function name is the substring before '('
        return formula.Substring(0, openParenIndex).Trim();
    }
}
