// Title: Set default row height and auto‑fit rows containing formulas with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet default row height and auto‑fit rows that have formulas | How to auto‑fit only rows that contain a formula using Aspose.Cells .NET | Iterate rows in a workbook to adjust height based on formula cells with Aspose.Cells | Detect formula cells and apply AutoFitRow after loading an Excel file in C# | Aspose.Cells .NET set StandardHeight then AutoFitRow for rows with formulas
// Tags: standardheight property Aspose.Cells | autofitrow based on formula detection | formula cell detection Aspose.Cells C# | worksheet row height adjustment Aspose.Cells | excel workbook processing Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing Excel workbook, sets each worksheet's StandardHeight to 15 points, scans every row for cells that contain formulas, auto‑fits only those rows, and saves the modified file, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Set the default row height (15 points)
                    sheet.Cells.StandardHeight = 15;

                    // Determine the last used row and column indices
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Scan rows for formulas
                    for (int row = 0; row <= maxRow; row++)
                    {
                        bool hasFormula = false;

                        // Check each cell in the current row for a formula
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                hasFormula = true;
                                break; // No need to check further cells in this row
                            }
                        }

                        // Auto‑fit the row height if a formula is present
                        if (hasFormula)
                        {
                            sheet.AutoFitRow(row);
                        }
                    }
                }
                catch (Exception sheetEx)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {sheetEx.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
