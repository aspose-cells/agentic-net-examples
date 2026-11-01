// Title: Convert multiple Excel (.xlsx/.xls) files to CSV in C# while trimming leading blank rows and preserving the first data row as header using Aspose.Cells
// Common Searches: C# Aspose.Cells batch convert .xlsx and .xls files to CSV with top empty rows removed | How to delete leading empty rows in the first worksheet before saving as CSV using Aspose.Cells | Preserve the first data row as header when exporting Excel workbooks to CSV in .NET | Iterate through a directory of Excel files and generate corresponding CSV files with Aspose.Cells | Trim blank rows at the beginning of a worksheet programmatically with Aspose.Cells API
// Tags: batch Excel to CSV conversion Aspose.Cells | remove top empty rows worksheet Aspose.Cells | header row preservation CSV export Aspose.Cells | process directory of .xlsx files C# | delete empty rows using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified folder for .xlsx and .xls workbooks, loads each with Aspose.Cells, removes any completely empty rows at the top of the first worksheet while keeping the first non‑empty row as the header, and saves the result as a CSV file with the same name in an output directory. It handles missing files and runtime errors gracefully.
class ExcelToCsvBatchConverter
{
    static void Main(string[] args)
    {
        try
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\InputExcels";
            // Folder where CSV files will be saved
            string outputFolder = @"C:\OutputCsvs";

            // Validate source folder
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each .xlsx or .xls file in the source folder
            foreach (string excelPath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                string ext = Path.GetExtension(excelPath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls")
                    continue; // Skip non‑Excel files

                // Verify the file exists before loading
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook (lifecycle rule: load)
                    Workbook workbook = new Workbook(excelPath);

                    // Work with the first worksheet (adjust if needed)
                    Worksheet sheet = workbook.Worksheets[0];

                    // Trim leading blank rows while preserving the first non‑empty row as header
                    TrimLeadingBlankRows(sheet);

                    // Build CSV output path (same file name, .csv extension)
                    string csvFileName = Path.GetFileNameWithoutExtension(excelPath) + ".csv";
                    string csvPath = Path.Combine(outputFolder, csvFileName);

                    // Save as CSV (lifecycle rule: save)
                    workbook.Save(csvPath, SaveFormat.Csv);
                    Console.WriteLine($"Converted: {excelPath} -> {csvPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Removes rows that are completely empty at the top of the worksheet.
    // The first non‑empty row becomes the header row and is kept.
    private static void TrimLeadingBlankRows(Worksheet sheet)
    {
        Cells cells = sheet.Cells;
        int firstNonEmptyRow = -1;

        // Scan rows from the top until a row with any non‑empty cell is found
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            bool rowHasData = false;
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                // A cell is considered non‑empty if its type is not Null
                if (cells[row, col].Type != CellValueType.IsNull)
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

        // If there are leading blank rows, delete them
        if (firstNonEmptyRow > 0)
        {
            // Delete rows from 0 to firstNonEmptyRow‑1
            sheet.Cells.DeleteRows(0, firstNonEmptyRow);
        }
    }
}
