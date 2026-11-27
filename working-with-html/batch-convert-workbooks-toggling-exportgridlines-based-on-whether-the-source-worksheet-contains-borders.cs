// Title: Batch convert Excel workbooks to HTML with conditional ExportGridLines based on worksheet border detection using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells batch convert xlsx to html exportgridlines only when borders are present | how to detect borders in an Excel worksheet with Aspose.Cells before saving as HTML | save multiple Excel files as HTML using Aspose.Cells and preserve hidden worksheets | set HtmlSaveOptions.ExportGridLines dynamically based on cell border style in C# | convert all workbooks in a folder to HTML with Aspose.Cells and conditional grid lines
// Tags: batch excel to html conversion Aspose.Cells | conditional ExportGridLines HtmlSaveOptions | worksheet border detection Aspose.Cells | preserve hidden worksheets Aspose.Cells HTML export | export entire workbook Aspose.Cells HtmlSaveOptions

using System;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookConversion
{
    // The program scans an input folder for Excel files, loads each workbook with Aspose.Cells, checks all worksheets for any visible cell borders, sets HtmlSaveOptions.ExportGridLines to true only when borders are found, and saves each workbook as an HTML file in an output folder, preserving hidden sheets and exporting the entire workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output directories (modify as needed or pass via args).
                string inputFolder = @"InputWorkbooks";
                string outputFolder = @"OutputHtml";

                // Verify input folder exists.
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder '{inputFolder}' does not exist.");
                    return;
                }

                // Ensure output folder exists.
                Directory.CreateDirectory(outputFolder);

                // Retrieve all files in the input folder.
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
                if (excelFiles.Length == 0)
                {
                    Console.WriteLine("No files found in the input folder.");
                    return;
                }

                foreach (string filePath in excelFiles)
                {
                    // Process only supported Excel extensions.
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                    {
                        Console.WriteLine($"Skipping unsupported file: {Path.GetFileName(filePath)}");
                        continue;
                    }

                    // Ensure the file actually exists before attempting to load.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        Console.WriteLine($"Processing: {Path.GetFileName(filePath)}");

                        // Load the workbook.
                        Workbook workbook = new Workbook(filePath);

                        // Determine whether any worksheet contains borders.
                        bool workbookHasBorders = false;
                        foreach (Worksheet ws in workbook.Worksheets)
                        {
                            if (WorksheetContainsBorders(ws))
                            {
                                workbookHasBorders = true;
                                break;
                            }
                        }

                        // Configure HTML save options.
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            // Export grid lines only when borders are present.
                            ExportGridLines = workbookHasBorders,

                            // Export the whole workbook (not only the active sheet).
                            ExportActiveWorksheetOnly = false,

                            // Keep hidden worksheets in the output.
                            ExportHiddenWorksheet = true
                        };

                        // Build output file name.
                        string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the workbook as HTML.
                        workbook.Save(outputPath, htmlOptions);

                        Console.WriteLine($"Saved HTML: {outputFileName} (ExportGridLines = {htmlOptions.ExportGridLines})");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <param name="worksheet">Worksheet to examine.</param>
        /// <returns>True if a border is found; otherwise false.</returns>
        private static bool WorksheetContainsBorders(Worksheet worksheet)
        {
            // Determine the used range to limit iteration.
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            // Empty sheet – no borders.
            if (maxRow < 0 || maxCol < 0)
                return false;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    Style style = cell.GetStyle();

                    // Examine each border side using the correct BorderType members.
                    if (HasVisibleBorder(style.Borders[BorderType.LeftBorder]) ||
                        HasVisibleBorder(style.Borders[BorderType.RightBorder]) ||
                        HasVisibleBorder(style.Borders[BorderType.TopBorder]) ||
                        HasVisibleBorder(style.Borders[BorderType.BottomBorder]) ||
                        HasVisibleBorder(style.Borders[BorderType.DiagonalDown]) ||
                        HasVisibleBorder(style.Borders[BorderType.DiagonalUp]) ||
                        HasVisibleBorder(style.Borders[BorderType.Vertical]) ||
                        HasVisibleBorder(style.Borders[BorderType.Horizontal]))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Determines whether the specified border is visible.
        /// </summary>
        /// <param name="border">Border to evaluate.</param>
        /// <returns>True if the border is visible; otherwise false.</returns>
        private static bool HasVisibleBorder(Border border)
        {
            return border != null && border.LineStyle != CellBorderType.None;
        }
    }
}
