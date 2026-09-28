// Title: Batch apply a holiday-themed cell style to date cells across all worksheets in multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads every .xlsx file from a given directory, creates a reusable holiday style, formats all DateTime cells in each worksheet with a light‑yellow background and red bold font using Aspose.Cells, and writes the updated workbooks to a separate output folder while logging errors. | Provide a C# example that recursively scans a folder for Excel workbooks, applies a custom holiday cell style to calendar entries, and saves each modified file, demonstrating proper resource handling with Aspose.Cells.
// Common Searches: asp.net batch process to highlight calendar dates in multiple Excel files with Aspose.Cells | c# script for applying holiday formatting to date cells across all worksheets in a folder | how to detect DateTime cells and set custom style using Aspose.Cells in bulk | automate theming of Excel workbooks for holiday season in .NET
// Tags: batch apply cell style Aspose.Cells .xlsx | holiday theme cell style C# | iterate worksheets apply style Aspose.Cells | load and save multiple workbooks Aspose.Cells | detect DateTime cell type Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid conflict with System.Range introduced in newer .NET versions
using CellsRange = Aspose.Cells.Range;

// The program scans all .xlsx files in a source directory, loads each workbook with Aspose.Cells, iterates through every worksheet, identifies cells of type DateTime, applies a light‑yellow background with red bold font as a holiday theme, and saves the modified workbooks to an output folder while handling missing files and logging errors.
class HolidayThemeBatch
{
    static void Main()
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\Workbooks\Input";

        // Folder where the themed workbooks will be saved
        string outputFolder = @"C:\Workbooks\Output";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        string[] workbookFiles;
        try
        {
            // Retrieve all Excel files from the input folder (including subfolders)
            workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing input folder: {ex.Message}");
            return;
        }

        foreach (string workbookPath in workbookFiles)
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                continue;
            }

            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(workbookPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Define a holiday style (e.g., light yellow fill with red bold font)
                    Style holidayStyle = workbook.CreateStyle();
                    holidayStyle.ForegroundColor = Color.LightYellow;
                    holidayStyle.Pattern = BackgroundType.Solid;
                    holidayStyle.Font.Color = Color.Red;
                    holidayStyle.Font.IsBold = true;

                    // Determine the used range of the worksheet
                    CellsRange usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange == null)
                        continue; // No data on the sheet

                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Scan all cells within the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Identify cells that contain DateTime values (considered as calendar entries)
                            if (cell.Type == CellValueType.IsDateTime)
                            {
                                // Apply the holiday style to the cell
                                cell.SetStyle(holidayStyle);
                            }
                        }
                    }
                }

                // Save the modified workbook to the output folder (lifecycle rule: save)
                string fileName = Path.GetFileName(workbookPath);
                string outputPath = Path.Combine(outputFolder, fileName);
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any errors but continue processing other files
                Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
            }
        }
    }
}
