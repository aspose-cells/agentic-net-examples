// Title: Convert an HTML file to an Excel workbook and generate in‑cell drop‑down lists from <select> elements using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load HTML file and add list validation from select tags | How to extract option values from HTML and create Excel data‑validation dropdowns with Aspose.Cells | Convert input.html to output.xlsx with in‑cell dropdowns using Aspose.Cells .NET | C# generate Excel list validation from HTML <select> elements via Aspose.Cells | Parse HTML select elements and apply Excel data validation using Aspose.Cells
// Tags: Aspose.Cells HTML to XLSX conversion | Aspose.Cells create list‑type data validation | Excel dropdown from HTML select using Aspose.Cells | HtmlLoadOptions workbook loading C# | Regex extraction of option values for Excel validation

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example reads an HTML file, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, uses regular expressions to locate each <select> element and collect its <option> texts, writes a placeholder into successive rows, and adds a list‑type data validation (in‑cell drop‑down) populated with the extracted values. The workbook is then saved as an XLSX file, with error handling for missing files and load/save failures.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source HTML file
                string htmlPath = "input.html";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                    return;
                }

                // Read the entire HTML content
                string html = File.ReadAllText(htmlPath);

                // ---------- Load HTML into workbook ----------
                Workbook workbook;
                try
                {
                    var loadOptions = new HtmlLoadOptions();
                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
                    {
                        // Use the Workbook constructor that accepts a stream and load options
                        workbook = new Workbook(ms, loadOptions);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load HTML into workbook: {ex.Message}");
                    return;
                }

                // ---------- Parse HTML to locate <select> elements ----------
                var selectMatches = Regex.Matches(html, @"<select[^>]*?>(.*?)</select>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                if (selectMatches.Count > 0)
                {
                    // Work with the first worksheet (adjust if needed)
                    Worksheet sheet = workbook.Worksheets[0];

                    // Starting row for placing validation cells (0‑based index)
                    int currentRow = 0;

                    foreach (Match selectMatch in selectMatches)
                    {
                        // Extract option texts within the current <select>
                        var optionMatches = Regex.Matches(selectMatch.Groups[1].Value, @"<option[^>]*?>(.*?)</option>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                        List<string> items = new List<string>();

                        foreach (Match optionMatch in optionMatches)
                        {
                            string text = optionMatch.Groups[1].Value.Trim();
                            if (!string.IsNullOrEmpty(text))
                                items.Add(text);
                        }

                        // Place a placeholder value in the target cell (column A)
                        Cell targetCell = sheet.Cells[currentRow, 0];
                        targetCell.PutValue("Select");

                        // Define the cell area for validation (single cell)
                        CellArea area = new CellArea
                        {
                            StartRow = currentRow,
                            StartColumn = 0,
                            EndRow = currentRow,
                            EndColumn = 0
                        };

                        // ---------- Create data‑validation list ----------
                        int validationIndex = sheet.Validations.Add(area);
                        Validation validation = sheet.Validations[validationIndex];
                        validation.Type = ValidationType.List;
                        validation.Operator = OperatorType.None;
                        validation.InCellDropDown = true;
                        validation.ShowError = true;
                        validation.ErrorTitle = "Invalid selection";
                        validation.ErrorMessage = "Please choose a value from the list.";
                        validation.AlertStyle = ValidationAlertType.Stop;

                        // Build the comma‑separated list for the formula
                        validation.Formula1 = string.Join(",", items);

                        // Move to the next row for the next <select>
                        currentRow++;
                    }
                }

                // ---------- Save workbook ----------
                string outputPath = "output.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
