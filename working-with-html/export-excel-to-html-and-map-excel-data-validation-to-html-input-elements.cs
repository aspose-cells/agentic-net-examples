// Title: Convert an Excel worksheet to an HTML table with data‑validation rendered as form controls using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel sheet to HTML with dropdowns for list validation | How to map Excel data validation to HTML input elements using Aspose.Cells in .NET | Generate HTML table from Excel file while preserving validation rules with Aspose.Cells C# | Aspose.Cells create HTML select element from Excel validation list in C# | Export Excel to HTML and convert numeric validation to <input type="number"> using Aspose.Cells
// Tags: Aspose.Cells HTML export with validation mapping | Excel list validation to HTML select Aspose.Cells | C# numeric validation to HTML number input | Excel date validation to HTML date field | Generate HTML table from Excel using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using Aspose.Cells;

namespace ExcelToHtmlWithValidation
{
    // The C# program loads an Excel workbook with Aspose.Cells, builds a lookup of each cell's Validation object, and iterates over the used range to produce an HTML table. Cells with validation are rendered as appropriate form controls—<select> for list validation, <input type="number"> with min/max for numeric validation, <input type="date"> for date validation, and <input type="text"> for other types—while all other cells are displayed as plain text. The resulting HTML file is saved to disk.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file
                string excelPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Input file \"{excelPath}\" not found.");
                    return;
                }

                // Path to the output HTML file
                string htmlPath = "output.html";

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(excelPath);

                // Get the first worksheet (adjust as needed)
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Build a lookup for cell address -> Validation
                Dictionary<string, Validation> validationMap = new Dictionary<string, Validation>(StringComparer.OrdinalIgnoreCase);
                foreach (Validation val in sheet.Validations)
                {
                    // A Validation can apply to a range; map each cell in the range
                    foreach (CellArea area in val.Areas)
                    {
                        for (int row = area.StartRow; row <= area.EndRow; row++)
                        {
                            for (int col = area.StartColumn; col <= area.EndColumn; col++)
                            {
                                string address = CellsHelper.CellIndexToName(row, col);
                                validationMap[address] = val;
                            }
                        }
                    }
                }

                // Build HTML output
                StringBuilder html = new StringBuilder();

                // Basic HTML header
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html>");
                html.AppendLine("<head>");
                html.AppendLine("<meta charset=\"UTF-8\">");
                html.AppendLine("<title>Excel Export</title>");
                html.AppendLine("<style>");
                html.AppendLine("table { border-collapse: collapse; }");
                html.AppendLine("td, th { border: 1px solid #999; padding: 4px; }");
                html.AppendLine("</style>");
                html.AppendLine("</head>");
                html.AppendLine("<body>");
                html.AppendLine("<table>");

                // Determine the used range
                int minRow = cells.MinRow;
                int maxRow = cells.MaxRow;
                int minCol = cells.MinColumn;
                int maxCol = cells.MaxColumn;

                // Optional: render header row with column letters
                html.AppendLine("<tr>");
                html.AppendLine("<th></th>"); // top-left empty cell
                for (int col = minCol; col <= maxCol; col++)
                {
                    string colName = CellsHelper.ColumnIndexToName(col);
                    html.AppendFormat("<th>{0}</th>", colName);
                    html.AppendLine();
                }
                html.AppendLine("</tr>");

                // Iterate through rows
                for (int row = minRow; row <= maxRow; row++)
                {
                    html.AppendLine("<tr>");

                    // Row header with row number
                    html.AppendFormat("<th>{0}</th>", (row + 1));
                    html.AppendLine();

                    for (int col = minCol; col <= maxCol; col++)
                    {
                        string address = CellsHelper.CellIndexToName(row, col);
                        Cell cell = cells[row, col];
                        string cellText = cell.Value?.ToString() ?? string.Empty;

                        if (validationMap.TryGetValue(address, out Validation val))
                        {
                            // Map validation to appropriate HTML input element
                            string inputHtml = GenerateInputHtml(val, cellText);
                            html.AppendFormat("<td>{0}</td>", inputHtml);
                        }
                        else
                        {
                            // No validation – just display the cell value
                            html.AppendFormat("<td>{0}</td>", WebUtility.HtmlEncode(cellText));
                        }

                        html.AppendLine();
                    }

                    html.AppendLine("</tr>");
                }

                html.AppendLine("</table>");
                html.AppendLine("</body>");
                html.AppendLine("</html>");

                // Save the HTML (lifecycle rule: save)
                File.WriteAllText(htmlPath, html.ToString(), Encoding.UTF8);
                Console.WriteLine($"HTML file generated at \"{htmlPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static string GenerateInputHtml(Validation val, string defaultValue)
        {
            // Escape default value for HTML attribute
            string escapedValue = WebUtility.HtmlEncode(defaultValue);

            switch (val.Type)
            {
                case ValidationType.List:
                    // List validation – create a <select> element
                    List<string> items = new List<string>();
                    if (!string.IsNullOrEmpty(val.Formula1) && !val.Formula1.StartsWith("="))
                    {
                        // Comma‑separated list
                        string[] parts = val.Formula1.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string p in parts)
                            items.Add(p.Trim());
                    }

                    StringBuilder select = new StringBuilder();
                    select.Append("<select>");
                    foreach (string opt in items)
                    {
                        string selected = (opt == defaultValue) ? " selected" : "";
                        select.AppendFormat("<option value=\"{0}\"{1}>{0}</option>", WebUtility.HtmlEncode(opt), selected);
                    }
                    select.Append("</select>");
                    return select.ToString();

                case ValidationType.WholeNumber:
                case ValidationType.Decimal:
                    // Numeric validation – create <input type="number">
                    string minAttr = "";
                    string maxAttr = "";

                    // Handle common operators
                    if (val.Operator == OperatorType.Between && !string.IsNullOrEmpty(val.Formula1) && !string.IsNullOrEmpty(val.Formula2))
                    {
                        minAttr = $" min=\"{val.Formula1}\"";
                        maxAttr = $" max=\"{val.Formula2}\"";
                    }
                    else if (val.Operator == OperatorType.GreaterThan && !string.IsNullOrEmpty(val.Formula1))
                    {
                        minAttr = $" min=\"{val.Formula1}\"";
                    }
                    else if (val.Operator == OperatorType.LessThan && !string.IsNullOrEmpty(val.Formula1))
                    {
                        maxAttr = $" max=\"{val.Formula1}\"";
                    }

                    return $"<input type=\"number\" value=\"{escapedValue}\"{minAttr}{maxAttr} />";

                case ValidationType.Date:
                    // Date validation – create <input type="date">
                    return $"<input type=\"date\" value=\"{escapedValue}\" />";

                case ValidationType.TextLength:
                    // Text length – create <input type=\"text\">
                    return $"<input type=\"text\" value=\"{escapedValue}\" />";

                default:
                    // Fallback – plain text input
                    return $"<input type=\"text\" value=\"{escapedValue}\" />";
            }
        }
    }
}
