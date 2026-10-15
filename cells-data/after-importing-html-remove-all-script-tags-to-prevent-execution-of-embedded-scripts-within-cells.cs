// Title: Import an HTML file into an Aspose.Cells workbook and strip all <script> tags from cell values using C#
// AI Prompts: Load an HTML document into a Workbook, iterate over every string cell, and remove any <script>...</script> blocks with a case‑insensitive regex before saving as XLSX. | Validate the source HTML file’s existence, catch exceptions, and apply RegexOptions.Singleline|IgnoreCase to clean script elements in cell contents across all worksheets. | Update cells only when the cleaned text differs from the original, then export the sanitized workbook to a new XLSX file using Aspose.Cells.
// Common Searches: c# aspocells remove script tags after loading html file | how to sanitize html content in excel cells using aspocells c# | aspocells regex strip <script> elements from imported html workbook | remove javascript from cells when importing html with aspocells | aspocells load html and clean script tags before saving as xlsx
// Tags: remove script tags from HTML cells Aspose.Cells | sanitize imported HTML Aspose.Cells C# | regex based script removal Excel cells | load HTML workbook Aspose.Cells C# | export cleaned workbook to XLSX Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlScriptRemoval
{
    // The example loads an HTML file into an Aspose.Cells Workbook, scans each string cell for <script>...</script> blocks, removes them using a case‑insensitive regular expression, updates only modified cells, and saves the sanitized workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the input file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: Input file '{htmlPath}' not found.");
                return;
            }

            try
            {
                // Load the HTML content into a workbook (Aspose.Cells can directly load HTML files)
                Workbook workbook = new Workbook(htmlPath);

                // Regular expression to match <script> tags (case‑insensitive, multiline)
                string scriptPattern = @"<script.*?>.*?</script>";
                Regex scriptRegex = new Regex(scriptPattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range to limit iteration to populated cells
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only string cells that may contain HTML
                            if (cell.Type == CellValueType.IsString)
                            {
                                string originalText = cell.StringValue;

                                // Remove any <script>...</script> blocks
                                string cleanedText = scriptRegex.Replace(originalText, string.Empty);

                                // Update the cell only if changes were made
                                if (!string.Equals(originalText, cleanedText, StringComparison.Ordinal))
                                {
                                    cell.PutValue(cleanedText);
                                }
                            }
                        }
                    }
                }

                // Save the cleaned workbook to a new file (e.g., XLSX)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"HTML imported and script tags removed. Saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
