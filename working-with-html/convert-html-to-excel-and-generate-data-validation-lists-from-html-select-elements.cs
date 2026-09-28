// Title: Convert an HTML file to Excel and generate dropdown lists from <select> elements using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an HTML document into an Aspose.Cells Workbook, extracts each <select> element's option values, and adds a list‑type data validation to successive cells. | Demonstrate how to parse <select> tags with regular expressions in C# and use the extracted options to create Excel dropdown lists via Aspose.Cells validations. | Provide a complete example that saves the modified workbook as an XLSX file after adding data‑validation lists derived from HTML select elements.
// Common Searches: aspocells c# convert html to xlsx with dropdown validation from select tags | how to add list validation to cells after loading html with Aspose.Cells .NET | extract option values from html select and create excel dropdown using Aspose.Cells | load html string into Aspose.Cells workbook via MemoryStream c# example
// Tags: Aspose.Cells HTML to XLSX conversion | C# data validation list creation | extract select options with regex | apply list validation to Excel cells | load HTML via MemoryStream in Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

// The program reads an HTML file, loads it into an Aspose.Cells Workbook using a memory stream, extracts option texts from each <select> element with regular expressions, adds a list‑type data validation (dropdown) to consecutive cells in column A, and saves the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Ensure the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: File '{htmlPath}' not found.");
                return;
            }

            // Read the HTML content
            string html = File.ReadAllText(htmlPath);

            // -------------------------------------------------
            // Load HTML into an Aspose.Cells workbook
            // -------------------------------------------------
            Workbook workbook;
            // Load from HTML string via a memory stream
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(ms, loadOptions);
            }

            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Parse the HTML to find <select> elements and their <option> values
            // -------------------------------------------------
            var selectMatches = Regex.Matches(html, @"<select[^>]*>(.*?)</select>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (selectMatches.Count > 0)
            {
                int currentRow = 0; // start placing validations from the first row

                foreach (Match selectMatch in selectMatches)
                {
                    string selectInnerHtml = selectMatch.Groups[1].Value;

                    // Gather all option texts for the current <select>
                    List<string> optionTexts = new List<string>();
                    var optionMatches = Regex.Matches(selectInnerHtml, @"<option[^>]*>(.*?)</option>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    foreach (Match optionMatch in optionMatches)
                    {
                        string text = optionMatch.Groups[1].Value.Trim();
                        if (!string.IsNullOrEmpty(text))
                            optionTexts.Add(text);
                    }

                    // Build a comma‑separated list for the validation formula
                    string listFormula = "\"" + string.Join(",", optionTexts) + "\"";

                    // -------------------------------------------------
                    // Add a data‑validation list to a cell (column A, current row)
                    // -------------------------------------------------
                    Cell targetCell = sheet.Cells[currentRow, 0]; // A1, A2, ...

                    // Define the cell area for the validation (single cell)
                    CellArea area = new CellArea
                    {
                        StartRow = currentRow,
                        EndRow = currentRow,
                        StartColumn = 0,
                        EndColumn = 0
                    };

                    int validationIndex = sheet.Validations.Add(area);
                    Validation validation = sheet.Validations[validationIndex];
                    validation.Type = ValidationType.List;
                    validation.Operator = OperatorType.None;
                    validation.Formula1 = listFormula;
                    validation.ShowError = true;
                    validation.ErrorTitle = "Invalid selection";
                    validation.ErrorMessage = "Please select a value from the list.";

                    currentRow++; // move to the next row for the next <select>
                }
            }

            // -------------------------------------------------
            // Save the workbook to an Excel file
            // -------------------------------------------------
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
