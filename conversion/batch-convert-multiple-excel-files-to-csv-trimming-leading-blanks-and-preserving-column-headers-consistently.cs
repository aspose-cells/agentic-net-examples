// Title: Batch convert .xls and .xlsx workbooks to CSV in C# with Aspose.Cells, removing leading blank rows and preserving the header row
// Common Searches: C# Aspose.Cells batch convert multiple Excel files to CSV | How to skip empty rows at top of Excel when exporting to CSV using Aspose.Cells | Preserve first row as header while converting .xls to CSV in .NET | Delete leading blank rows before saving workbook as CSV with Aspose.Cells | Convert all Excel workbooks in a folder to CSV using Aspose.Cells for .NET
// Tags: batch Excel to CSV conversion Aspose.Cells | trim leading empty rows Aspose.Cells | keep header row when exporting CSV .NET | convert multiple .xls and .xlsx files programmatically | delete rows before CSV save Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace ExcelConversion
{
    // The C# program scans a specified input directory for .xls and .xlsx files, loads each workbook with Aspose.Cells, removes any leading rows that contain only empty cells, and saves the first worksheet as a CSV file in an output directory, ensuring the first non‑empty row (the header) is retained.
    class ExcelToCsvBatchConverter
    {
        static void Main()
        {
            // Folder containing source Excel files
            string inputFolder = @"C:\InputExcelFiles";
            // Folder where CSV files will be saved
            string outputFolder = @"C:\OutputCsvFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Gather .xlsx and .xls files
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");
            excelFiles = excelFiles.Concat(Directory.GetFiles(inputFolder, "*.xls")).ToArray();

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found, skipping: {excelPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Work with the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Find the first row that contains any non‑empty cell
                    int firstNonEmptyRow = -1;
                    for (int row = 0; row <= cells.MaxDataRow; row++)
                    {
                        bool rowHasData = false;
                        for (int col = 0; col <= cells.MaxDataColumn; col++)
                        {
                            if (!string.IsNullOrWhiteSpace(cells[row, col].StringValue))
                            {
                                rowHasData = true;
                                break;
                            }
                        }
                        if (rowHasData)
                        {
                            firstNonEmptyRow = row;
                            break;
                        }
                    }

                    // Delete leading blank rows if needed
                    if (firstNonEmptyRow > 0)
                    {
                        cells.DeleteRows(0, firstNonEmptyRow);
                    }

                    // Build output CSV file path
                    string csvFileName = Path.GetFileNameWithoutExtension(excelPath) + ".csv";
                    string csvPath = Path.Combine(outputFolder, csvFileName);

                    // Save as CSV
                    workbook.Save(csvPath, SaveFormat.Csv);
                    Console.WriteLine($"Converted: {Path.GetFileName(excelPath)} -> {csvFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
