// Title: Convert an Excel workbook to HTML and map Aspose.Cells data‑validation rules to corresponding HTML form controls (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates through each worksheet, and generates an HTML file where cells with data‑validation are rendered as appropriate <input> or <select> elements. | Enhance the generated HTML by adding CSS classes that indicate the validation type (list, numeric, date, text length) and ensure all worksheets in the workbook are processed. | Add support for list‑type validations that reference a range or named range, extracting the allowed values and populating a <select> dropdown in the output HTML.
// Common Searches: how to export Excel to HTML with Aspose.Cells and preserve data validation as form fields in C# | C# Aspose.Cells generate HTML table with input elements for cells that have validation rules | map Excel list validation to HTML select using Aspose.Cells .NET | render numeric and date validation from Excel as HTML input type number/date with Aspose.Cells | extract range‑based validation options from an Excel worksheet when converting to HTML
// Tags: Aspose.Cells HTML export with data validation | Excel list validation to HTML select C# | numeric validation to number input Aspose.Cells | date validation to date input Aspose.Cells | extract validation options from range Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Net;
using Aspose.Cells;

// The program loads an Excel workbook using Aspose.Cells, builds a lookup of cell addresses to their Validation objects, and writes an HTML file where each cell is displayed either as plain text or as an appropriate <input> or <select> element based on its validation type (list, numeric, date, text length, etc.).
class ExcelToHtmlWithValidation
{
    static void Main()
    {
        try
        {
            // Load the Excel workbook (use the provided load rule)
            string excelPath = "input.xlsx";

            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            Workbook workbook = new Workbook(excelPath);

            // Prepare a StringBuilder to hold the generated HTML
            StringBuilder html = new StringBuilder();

            // Basic HTML header
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("<meta charset=\"UTF-8\">");
            html.AppendLine("<title>Excel Export</title>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                html.AppendLine($"<h2>{WebUtility.HtmlEncode(sheet.Name)}</h2>");
                html.AppendLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">");

                // Build a lookup for data validations for quick access
                var validationMap = BuildValidationMap(sheet);

                // Determine the used range of the sheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Iterate through rows
                for (int row = 0; row <= maxRow; row++)
                {
                    html.AppendLine("<tr>");
                    // Iterate through columns
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        string cellAddress = cell.Name; // e.g., "A1"

                        // Check if this cell has a data validation rule
                        if (validationMap.TryGetValue(cellAddress, out Validation dv))
                        {
                            // Render an appropriate HTML input element based on the validation type
                            string inputHtml = RenderInputForValidation(cell, dv);
                            html.AppendLine($"<td>{inputHtml}</td>");
                        }
                        else
                        {
                            // No validation – just output the cell's displayed value
                            string displayedValue = cell.StringValue ?? string.Empty;
                            displayedValue = WebUtility.HtmlEncode(displayedValue);
                            html.AppendLine($"<td>{displayedValue}</td>");
                        }
                    }
                    html.AppendLine("</tr>");
                }

                html.AppendLine("</table>");
            }

            // Close HTML tags
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            // Save the HTML content to a file (use the provided save rule)
            string htmlPath = "output.html";
            File.WriteAllText(htmlPath, html.ToString(), Encoding.UTF8);
            Console.WriteLine($"HTML file generated successfully at \"{htmlPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Builds a dictionary that maps cell addresses to their Validation objects
    private static Dictionary<string, Validation> BuildValidationMap(Worksheet sheet)
    {
        var map = new Dictionary<string, Validation>(StringComparer.OrdinalIgnoreCase);
        foreach (Validation dv in sheet.Validations)
        {
            // A validation can apply to a range of cells; iterate through each cell in the range
            foreach (CellArea area in dv.Areas)
            {
                for (int r = area.StartRow; r <= area.EndRow; r++)
                {
                    for (int c = area.StartColumn; c <= area.EndColumn; c++)
                    {
                        string address = CellsHelper.CellIndexToName(r, c);
                        // If multiple validations overlap, keep the first one encountered
                        if (!map.ContainsKey(address))
                            map[address] = dv;
                    }
                }
            }
        }
        return map;
    }

    // Generates HTML input/select markup based on the type of Validation
    private static string RenderInputForValidation(Cell cell, Validation dv)
    {
        // Determine the validation type
        switch (dv.Type)
        {
            case ValidationType.List:
                // List validation – render a <select> element
                var options = GetListOptions(dv, cell.Worksheet);
                StringBuilder select = new StringBuilder();
                select.Append("<select>");
                foreach (string opt in options)
                {
                    string encoded = WebUtility.HtmlEncode(opt);
                    select.Append($"<option value=\"{encoded}\">{encoded}</option>");
                }
                select.Append("</select>");
                return select.ToString();

            case ValidationType.WholeNumber:
            case ValidationType.Decimal:
                // Numeric validation – render an <input type=\"number\">
                string numVal = WebUtility.HtmlEncode(cell.StringValue ?? string.Empty);
                return $"<input type=\"number\" value=\"{numVal}\" />";

            case ValidationType.Date:
                // Date validation – render an <input type=\"date\">
                string dateVal = cell.StringValue ?? string.Empty;
                if (DateTime.TryParse(dateVal, out DateTime dt))
                    dateVal = dt.ToString("yyyy-MM-dd");
                dateVal = WebUtility.HtmlEncode(dateVal);
                return $"<input type=\"date\" value=\"{dateVal}\" />";

            case ValidationType.TextLength:
                // Text length – render a simple text input
                string txt = WebUtility.HtmlEncode(cell.StringValue ?? string.Empty);
                return $"<input type=\"text\" value=\"{txt}\" />";

            default:
                // Fallback – plain text input
                string fallback = WebUtility.HtmlEncode(cell.StringValue ?? string.Empty);
                return $"<input type=\"text\" value=\"{fallback}\" />";
        }
    }

    // Retrieves the list of allowed values for a List-type validation
    private static List<string> GetListOptions(Validation dv, Worksheet sheet)
    {
        var options = new List<string>();

        // The list can be defined directly (comma‑separated) or via a range reference
        string formula = dv.Formula1?.Trim();

        if (string.IsNullOrEmpty(formula))
            return options;

        // If the formula starts with '=', it's a range reference
        if (formula.StartsWith("="))
        {
            // Remove leading '='
            string rangeRef = formula.Substring(1);
            try
            {
                // Resolve the range (supports same‑sheet references)
                Aspose.Cells.Range range = sheet.Cells.CreateRange(rangeRef);
                if (range != null)
                {
                    for (int r = 0; r < range.RowCount; r++)
                    {
                        for (int c = 0; c < range.ColumnCount; c++)
                        {
                            string val = range[r, c].StringValue ?? string.Empty;
                            options.Add(val);
                        }
                    }
                }
            }
            catch (Exception)
            {
                // If the range cannot be resolved, return empty options
            }
        }
        else
        {
            // Comma‑separated list
            string[] parts = formula.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
                options.Add(part.Trim());
        }

        return options;
    }
}
