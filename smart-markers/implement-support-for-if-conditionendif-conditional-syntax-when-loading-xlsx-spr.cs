using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsConditionalSyntaxDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook without parsing formulas on open.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Regular expression to capture {{IF condition}}truePart{{ENDIF}}
            Regex ifRegex = new Regex(@"\{\{IF\s+(?<cond>[^}]+)\}\}(?<truePart>.*?)\{\{ENDIF\}\}",
                                      RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells.
                foreach (Cell cell in cells)
                {
                    // Only process cells that contain a string value.
                    if (cell.Type == CellValueType.IsString)
                    {
                        string cellText = cell.StringValue;
                        Match match = ifRegex.Match(cellText);
                        if (match.Success)
                        {
                            // Extract condition and true part.
                            string condition = match.Groups["cond"].Value.Trim();
                            string truePart = match.Groups["truePart"].Value.Trim();

                            // Escape double quotes inside the true part.
                            string escapedTruePart = truePart.Replace("\"", "\"\"");

                            // Build the Excel IF formula.
                            string formula = $"=IF({condition},\"{escapedTruePart}\",\"\")";

                            // Preserve the original style.
                            Style originalStyle = cell.GetStyle();

                            // Set the formula.
                            cell.Formula = formula;

                            // Reapply the original style to keep formatting.
                            cell.SetStyle(originalStyle);
                        }
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save("output.xlsx");
        }
    }
}