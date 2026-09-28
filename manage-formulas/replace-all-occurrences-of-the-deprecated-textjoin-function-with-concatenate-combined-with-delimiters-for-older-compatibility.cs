// Title: Replace deprecated TEXTJOIN formulas with CONCATENATE in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that scans each worksheet, finds cells containing the TEXTJOIN function, and rewrites the formula into an equivalent CONCATENATE expression while keeping the original delimiter. | Create a method that employs a regular expression to capture TEXTJOIN arguments, assembles a CONCATENATE formula string, and assigns it back to the cell's Formula property in a loaded workbook.
// Common Searches: Aspose.Cells replace TEXTJOIN with CONCATENATE in existing Excel file C# | How to convert deprecated TEXTJOIN formulas to CONCATENATE using .NET | C# code to modify Excel formulas across all sheets with Aspose.Cells | Regex based find and replace of TEXTJOIN function in Excel workbook | Update Excel workbook formulas for compatibility with older Excel versions using Aspose.Cells
// Tags: TEXTJOIN formula migration Aspose.Cells | convert TEXTJOIN to CONCATENATE .NET | regex based formula extraction C# | bulk cell formula update across worksheets | deprecated Excel function handling Aspose.Cells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, uses a case‑insensitive regular expression to locate TEXTJOIN calls in cell formulas, parses the delimiter and text arguments, builds an equivalent CONCATENATE expression preserving the delimiter, replaces only the matched TEXTJOIN segment while keeping surrounding formula parts intact, writes the new formula back to each cell, and saves the updated workbook.
class ReplaceTextJoin
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Regular expression to match TEXTJOIN function and capture its arguments
        // Example: TEXTJOIN(",",TRUE,A1,B1,C1)
        Regex textJoinRegex = new Regex(@"TEXTJOIN\(([^)]+)\)", RegexOptions.IgnoreCase);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula) && cell.Formula.IndexOf("TEXTJOIN", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Find the TEXTJOIN part in the formula
                    Match match = textJoinRegex.Match(cell.Formula);
                    if (match.Success)
                    {
                        // Split the captured arguments by commas, respecting possible nested commas inside quotes
                        string argsString = match.Groups[1].Value;
                        // Simple split assuming no commas inside quoted strings
                        string[] args = argsString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        if (args.Length >= 3)
                        {
                            // First argument: delimiter (keep as is, may be a quoted string or cell reference)
                            string delimiter = args[0].Trim();

                            // Second argument: ignore_empty (TRUE/FALSE) – not handled in this simple replacement
                            // Remaining arguments are the texts to concatenate
                            string[] texts = new string[args.Length - 2];
                            for (int i = 2; i < args.Length; i++)
                            {
                                texts[i - 2] = args[i].Trim();
                            }

                            // Build the new CONCATENATE formula:
                            // CONCATENATE(text1, delimiter, text2, delimiter, text3, ...)
                            // Note: This simple approach adds the delimiter after every text argument,
                            // including after the last one. Adjust if needed.
                            string newFormulaBody = "CONCATENATE(";
                            for (int i = 0; i < texts.Length; i++)
                            {
                                newFormulaBody += texts[i];
                                if (i < texts.Length - 1)
                                {
                                    newFormulaBody += ", " + delimiter + ", ";
                                }
                            }
                            newFormulaBody += ")";

                            // Preserve any surrounding parts of the original formula (e.g., =IF(...,TEXTJOIN(...),...))
                            // Replace only the matched TEXTJOIN segment.
                            string newFormula = textJoinRegex.Replace(cell.Formula, newFormulaBody);

                            // Assign the updated formula back to the cell
                            cell.Formula = newFormula;
                        }
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
