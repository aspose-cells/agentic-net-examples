// Title: Convert HTML to Excel with Aspose.Cells and create dropdown validation from <select> tags (C#)
// Description: Loads an HTML file into an Aspose.Cells Workbook, extracts <select> elements (using the element's id as the target cell address), builds a comma‑separated list of option values, applies a list‑type data‑validation rule to the corresponding cells, and saves the result as an XLSX file with in‑cell dropdowns.
// Keywords: Aspose.Cells HTML to Excel | C# data validation list | Excel dropdown from HTML select | .NET parse select options | HTML table conversion Aspose | programmatic list validation | convert HTML form to Excel
// Common Searches: Aspose.Cells convert HTML to Excel with dropdowns | C# create Excel data validation from <select> tags | How to map HTML select id to Excel cell using Aspose | Generate Excel list validation from HTML form | Parse HTML select options in .NET for Excel
// Developer Intent: Transform an HTML document into an Excel workbook and turn each <select> element into a cell‑level dropdown list using Aspose.Cells.
// Use Cases: Build Excel templates from web‑form pages, preserving choice fields as validated dropdowns. | Migrate survey definitions written in HTML to Excel while keeping answer options as list validations. | Automate report pipelines that convert HTML tables to Excel and add selectable options for specific cells.
// AI Prompts: Generate C# code with Aspose.Cells that loads an HTML file, extracts <select> elements, and adds list validation to cells based on each element's id. | Rewrite the SimpleHtmlParser to use HtmlAgilityPack for robust <select> and <option> extraction while keeping the existing validation logic.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells Workbook, extracts <select> elements (using the element's id as the target cell address), builds a comma‑separated list of option values, applies a list‑type data‑validation rule to the corresponding cells, and saves the result as an XLSX file with in‑cell dropdowns.
class HtmlToExcelWithValidation
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // --------------------------------------------------------------------
            // 1. Load the HTML file into a Workbook using HtmlLoadOptions.
            // --------------------------------------------------------------------
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                // Enable conversion of HTML tables to ListObjects (optional)
                TableLoadOptions = { TableToListObject = false }
            };

            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // --------------------------------------------------------------------
            // 2. Parse the HTML to extract <select> elements and their options.
            //    For simplicity, this example assumes each <select> has an "id"
            //    attribute that matches the target cell address (e.g., id="B3").
            // --------------------------------------------------------------------
            string htmlContent = File.ReadAllText(htmlPath);
            List<SimpleHtmlParser.SelectInfo> selectElements = SimpleHtmlParser.ExtractSelectElements(htmlContent);

            // --------------------------------------------------------------------
            // 3. For each <select>, create a data‑validation list in the
            //    corresponding cell.
            // --------------------------------------------------------------------
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: No worksheets were loaded from the HTML.");
                return;
            }

            Worksheet sheet = workbook.Worksheets[0];

            foreach (var select in selectElements)
            {
                // The cell address is taken from the select's id attribute.
                string cellAddress = select.Id?.Trim();
                if (string.IsNullOrEmpty(cellAddress))
                    continue; // skip if no mapping

                // Build a comma‑separated list of option values.
                string listFormula = string.Join(",", select.Options);

                try
                {
                    // Resolve the cell to obtain row/column indices.
                    Cell cell = sheet.Cells[cellAddress];
                    CellArea area = new CellArea
                    {
                        StartRow = cell.Row,
                        StartColumn = cell.Column,
                        EndRow = cell.Row,
                        EndColumn = cell.Column
                    };

                    // Add a validation rule to the worksheet for the specific cell.
                    int validationIndex = sheet.Validations.Add(area);
                    Validation validation = sheet.Validations[validationIndex];
                    validation.Type = ValidationType.List;
                    validation.InCellDropDown = true;               // show dropdown arrow
                    validation.Formula1 = listFormula;              // e.g., "Red,Green,Blue"
                    validation.Operator = OperatorType.None;
                    validation.ShowError = true;
                    validation.ErrorTitle = "Invalid selection";
                    validation.ErrorMessage = "Please choose a value from the list.";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not apply validation to '{cellAddress}'. {ex.Message}");
                }
            }

            // --------------------------------------------------------------------
            // 4. Save the workbook as an Excel file.
            // --------------------------------------------------------------------
            string excelPath = "output.xlsx";
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML converted to Excel with validation lists: {excelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// ------------------------------------------------------------------------
// Minimal HTML parser for <select> elements.
// This is a lightweight implementation; for complex HTML use HtmlAgilityPack.
// ------------------------------------------------------------------------
static class SimpleHtmlParser
{
    public class SelectInfo
    {
        public string? Id { get; set; }                 // Expected to be a cell address like "C5"
        public List<string> Options { get; set; } = new List<string>();
    }

    public static List<SelectInfo> ExtractSelectElements(string html)
    {
        var result = new List<SelectInfo>();
        int pos = 0;

        while ((pos = html.IndexOf("<select", pos, StringComparison.OrdinalIgnoreCase)) != -1)
        {
            // Find the end of the opening tag
            int startTagEnd = html.IndexOf('>', pos);
            if (startTagEnd == -1) break;

            // Extract the opening tag substring
            string openingTag = html.Substring(pos, startTagEnd - pos + 1);

            // Get the id attribute (if present)
            string? id = GetAttributeValue(openingTag, "id");

            // Locate the closing </select>
            int closeTagStart = html.IndexOf("</select>", startTagEnd, StringComparison.OrdinalIgnoreCase);
            if (closeTagStart == -1) break;

            // Extract inner HTML of the select
            string innerHtml = html.Substring(startTagEnd + 1, closeTagStart - startTagEnd - 1);

            // Parse <option> values
            var options = new List<string>();
            int optPos = 0;
            while ((optPos = innerHtml.IndexOf("<option", optPos, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                int optStartEnd = innerHtml.IndexOf('>', optPos);
                if (optStartEnd == -1) break;
                int optClose = innerHtml.IndexOf("</option>", optStartEnd, StringComparison.OrdinalIgnoreCase);
                if (optClose == -1) break;

                string optionText = innerHtml.Substring(optStartEnd + 1, optClose - optStartEnd - 1).Trim();

                // If the option has a value attribute, prefer it
                string optionTag = innerHtml.Substring(optPos, optStartEnd - optPos + 1);
                string? valueAttr = GetAttributeValue(optionTag, "value");
                if (!string.IsNullOrEmpty(valueAttr))
                    optionText = valueAttr;

                if (!string.IsNullOrEmpty(optionText))
                    options.Add(optionText);

                optPos = optClose + 9; // length of "</option>"
            }

            result.Add(new SelectInfo { Id = id, Options = options });

            pos = closeTagStart + 9; // move past "</select>"
        }

        return result;
    }

    private static string? GetAttributeValue(string tag, string attributeName)
    {
        string search = attributeName + "=";
        int idx = tag.IndexOf(search, StringComparison.OrdinalIgnoreCase);
        if (idx == -1) return null;

        int valueStart = idx + search.Length;
        if (valueStart >= tag.Length) return null;

        char quote = tag[valueStart];
        if (quote != '\'' && quote != '\"')
            return null;

        int valueEnd = tag.IndexOf(quote, valueStart + 1);
        if (valueEnd == -1) return null;

        return tag.Substring(valueStart + 1, valueEnd - valueStart - 1);
    }
}
