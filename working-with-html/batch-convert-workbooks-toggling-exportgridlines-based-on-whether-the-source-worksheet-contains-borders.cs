// Title: Batch convert Excel to HTML with conditional grid lines – Aspose.Cells .NET
// Description: Scans a folder for .xlsx/.xls/.xlsm files, detects any cell borders in each workbook, and saves the files as HTML. HtmlSaveOptions.ExportGridLines is turned on only when borders are found, ensuring the generated pages keep the original visual layout while avoiding unnecessary grid lines.
// Keywords: Aspose.Cells batch conversion | Excel to HTML C# | ExportGridLines conditional | detect cell borders Aspose | HtmlSaveOptions grid lines | process multiple workbooks | .NET Excel HTML export
// Common Searches: Aspose.Cells export Excel to HTML with grid lines only when borders exist | batch convert Excel files to HTML C# Aspose | how to toggle ExportGridLines based on workbook borders | detect borders in Excel workbook using Aspose.Cells | save multiple workbooks as HTML with conditional grid lines
// Developer Intent: Automatically convert every Excel file in a directory to HTML, enabling grid lines only if any worksheet contains cell borders.
// Use Cases: Generate web‑ready reports from a library of spreadsheets while preserving border styling for readability. | Automate nightly conversion of uploaded Excel documents to HTML, showing grid lines only when borders are present to reduce visual clutter. | Create a command‑line utility that processes large batches of workbooks for publishing on intranet portals.
// AI Prompts: Write a C# function that returns true when any cell in an Aspose.Cells workbook has a non‑none border. | Provide code to batch convert .xlsx, .xls, and .xlsm files in a folder to HTML, setting HtmlSaveOptions.ExportGridLines based on border detection. | Explain the effect of HtmlSaveOptions.ExportGridLines on the HTML output when cell borders are present versus absent.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace BatchConvertWithGridLines
{
    // Scans a folder for .xlsx/.xls/.xlsm files, detects any cell borders in each workbook, and saves the files as HTML. HtmlSaveOptions.ExportGridLines is turned on only when borders are found, ensuring the generated pages keep the original visual layout while avoiding unnecessary grid lines.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Source folder containing Excel files
            string sourceFolder = @"C:\InputWorkbooks";
            // Destination folder for converted HTML files
            string destFolder = @"C:\OutputHtml";

            ConvertWorkbooks(sourceFolder, destFolder);
        }

        /// <param name="sourceFolder">Folder with source .xlsx/.xls/.xlsm files.</param>
        /// <param name="destFolder">Folder where HTML files will be saved.</param>
        static void ConvertWorkbooks(string sourceFolder, string destFolder)
        {
            // Ensure destination folder exists
            Directory.CreateDirectory(destFolder);

            string[] files;
            try
            {
                // Get all files in the source folder
                files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing source folder '{sourceFolder}': {ex.Message}");
                return;
            }

            foreach (string filePath in files)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                    continue; // Skip non‑Excel files

                // Verify the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Determine if any worksheet contains borders
                    bool hasBorders = WorkbookContainsBorders(workbook);

                    // Prepare HTML save options and set ExportGridLines accordingly
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExportGridLines = hasBorders,
                        ExportActiveWorksheetOnly = false // export all worksheets
                    };

                    // Build output file name (same base name with .html extension)
                    string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                    string outputPath = Path.Combine(destFolder, outputFileName);

                    // Save the workbook as HTML using the prepared options
                    workbook.Save(outputPath, htmlOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }

        /// <param name="workbook">The workbook to inspect.</param>
        /// <returns>True if a border is found; otherwise false.</returns>
        static bool WorkbookContainsBorders(Workbook workbook)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range to limit iteration
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                // If the sheet is empty, skip it
                if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                    continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        // Retrieve the style of the current cell
                        Style style = sheet.Cells[row, col].GetStyle();

                        // Check each border side for a non‑none line style
                        if (style.Borders[BorderType.TopBorder].LineStyle != CellBorderType.None ||
                            style.Borders[BorderType.BottomBorder].LineStyle != CellBorderType.None ||
                            style.Borders[BorderType.LeftBorder].LineStyle != CellBorderType.None ||
                            style.Borders[BorderType.RightBorder].LineStyle != CellBorderType.None)
                        {
                            return true; // Border found, no need to continue
                        }
                    }
                }
            }
            return false; // No borders detected in any worksheet
        }
    }
}
