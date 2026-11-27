// Title: Export all worksheet formulas with addresses and last evaluated values to a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export formula cells with addresses and values to CSV | How to list Excel formulas and their evaluated results using Aspose.Cells .NET | Generate CSV of formula address, expression, and result from an .xlsx file in C# | Retrieve cell formula and last calculated value with Aspose.Cells and save to CSV | C# code to iterate worksheets and write formulas to a CSV using Aspose.Cells
// Tags: Aspose.Cells export formula cells to CSV | .NET retrieve formula address and value | C# enumerate formula cells across worksheets | CSV output of Excel formulas Aspose.Cells | last calculated value extraction .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through each worksheet and every cell containing a formula, captures the cell's A1-style address (including sheet name), the formula text, and the last evaluated value, escapes these fields for CSV compliance, and writes the collected data to a UTF‑8 encoded CSV file.
class ExportFormulasToCsv
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Prepare a StringBuilder to build CSV content
        StringBuilder csvBuilder = new StringBuilder();

        // Write CSV header
        csvBuilder.AppendLine("Address,Formula,LastEvaluatedValue");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection of the current worksheet
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Cell address in A1 style (including sheet name)
                    string address = $"{sheet.Name}!{cell.Name}";

                    // Formula text as stored in the cell
                    string formula = cell.Formula;

                    // Last evaluated value (as string, handling nulls)
                    string value = cell.Value?.ToString() ?? string.Empty;

                    // Escape commas and quotes for CSV compliance
                    address = $"\"{address.Replace("\"", "\"\"")}\"";
                    formula = $"\"{formula.Replace("\"", "\"\"")}\"";
                    value = $"\"{value.Replace("\"", "\"\"")}\"";

                    // Append the CSV line
                    csvBuilder.AppendLine($"{address},{formula},{value}");
                }
            }
        }

        // Save the CSV content to a file (replace with desired output path)
        File.WriteAllText("formulas.csv", csvBuilder.ToString(), Encoding.UTF8);
    }
}
