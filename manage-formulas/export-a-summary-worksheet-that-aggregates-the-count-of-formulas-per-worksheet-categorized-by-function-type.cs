using System;
using System.Collections.Generic;
using Aspose.Cells;
using System.Text.RegularExpressions;

class FormulaSummaryExporter
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output_with_summary.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Dictionary to hold counts: Worksheet name -> (Function name -> Count)
        var summary = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        // Iterate through each worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Prepare inner dictionary for this worksheet
            var funcCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            summary[ws.Name] = funcCounts;

            // Iterate through all cells that contain data/formulas
            foreach (Cell cell in ws.Cells)
            {
                // Check if the cell has a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    string functionName = ExtractFunctionName(cell.Formula);
                    if (!string.IsNullOrEmpty(functionName))
                    {
                        if (funcCounts.ContainsKey(functionName))
                            funcCounts[functionName]++;
                        else
                            funcCounts[functionName] = 1;
                    }
                }
            }
        }

        // Add a new worksheet for the summary
        int summaryIndex = workbook.Worksheets.Add();
        Worksheet summaryWs = workbook.Worksheets[summaryIndex];
        summaryWs.Name = "Summary";

        // Write headers
        summaryWs.Cells[0, 0].PutValue("Worksheet");
        summaryWs.Cells[0, 1].PutValue("Function");
        summaryWs.Cells[0, 2].PutValue("Count");

        // Populate summary data
        int row = 1;
        foreach (var wsEntry in summary)
        {
            string wsName = wsEntry.Key;
            foreach (var funcEntry in wsEntry.Value)
            {
                summaryWs.Cells[row, 0].PutValue(wsName);
                summaryWs.Cells[row, 1].PutValue(funcEntry.Key);
                summaryWs.Cells[row, 2].PutValue(funcEntry.Value);
                row++;
            }
        }

        // Save the workbook with the new summary sheet
        workbook.Save(outputPath);
    }

    // Helper method to extract the function name from a formula string
    private static string ExtractFunctionName(string formula)
    {
        // Remove leading '=' and any surrounding braces for array formulas
        string trimmed = formula.Trim();
        if (trimmed.StartsWith("="))
            trimmed = trimmed.Substring(1).Trim();

        // Remove leading '{' and trailing '}' for array formulas
        if (trimmed.StartsWith("{") && trimmed.EndsWith("}"))
            trimmed = trimmed.Substring(1, trimmed.Length - 2).Trim();

        // Use regex to capture the function name before the first '('
        Match match = Regex.Match(trimmed, @"^([A-Za-z_][A-Za-z0-9_]*)\s*\(");
        if (match.Success)
            return match.Groups[1].Value.ToUpperInvariant();

        // If no '(' is found, it might be a reference or named range; ignore it
        return null;
    }
}