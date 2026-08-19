// Title: Batch convert Excel files to PDF with automatic Timeline control using Aspose.Cells for .NET
// Description: Iterates over all Excel workbooks in a folder, detects the first DateTime column, creates a PivotTable, inserts a linked Timeline control at E1, saves a temporary XLSX, converts it to PDF with ConversionUtility, and writes the PDF to an output directory while cleaning up temporary files.
// Keywords: Aspose.Cells timeline | C# batch Excel to PDF | add timeline to workbook | programmatic pivot table Aspose | convert multiple Excel files to PDF | Excel automation .NET | timeline control C# | bulk Excel PDF conversion
// Common Searches: Aspose.Cells add timeline to each Excel file in a folder | C# batch convert Excel to PDF with timeline control | automate pivot table and timeline creation for Excel PDFs | how to insert timeline in multiple workbooks using Aspose | bulk Excel to PDF conversion with date filter
// Developer Intent: Automatically insert a Timeline control linked to a PivotTable in every Excel workbook of a directory and then convert each updated workbook to PDF in a single batch operation.
// Use Cases: Generate PDF reports for a set of financial spreadsheets by adding a date‑driven timeline and exporting each file. | Automate monthly sales dashboards: create a pivot table, attach a timeline for date filtering, and produce PDFs for stakeholder distribution. | Integrate into CI/CD pipelines to validate incoming Excel data, embed interactive timelines, and archive the results as PDFs.
// AI Prompts: Write C# code that scans a directory for Excel files, adds a timeline linked to a pivot table using Aspose.Cells, and saves each workbook as a PDF. | Explain how to move the timeline from cell E1 to a custom address and customize its style in the batch conversion script. | Suggest improvements for error handling when a date column is missing or PDF conversion fails in the provided batch process.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Utility;

// Iterates over all Excel workbooks in a folder, detects the first DateTime column, creates a PivotTable, inserts a linked Timeline control at E1, saves a temporary XLSX, converts it to PDF with ConversionUtility, and writes the PDF to an output directory while cleaning up temporary files.
public class BatchExcelToPdfWithTimeline
{
    // Entry point for the batch process
    public static void Run(string inputDirectory, string outputDirectory)
    {
        // Ensure input and output directories exist
        if (!Directory.Exists(inputDirectory))
            throw new DirectoryNotFoundException($"Input directory not found: {inputDirectory}");

        Directory.CreateDirectory(outputDirectory);

        // Process each Excel file in the input folder (supports .xlsx, .xls, .xlsm, .xlsb)
        foreach (string sourcePath in Directory.GetFiles(inputDirectory, "*.*", SearchOption.TopDirectoryOnly))
        {
            // Filter only supported Excel formats based on extension
            string ext = Path.GetExtension(sourcePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                continue; // skip non‑Excel files

            // Verify the file still exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"File not found: {sourcePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Locate a column that contains DateTime values to use as the timeline base field
                // ------------------------------------------------------------
                int dateColumnIndex = -1;
                int maxRows = sheet.Cells.MaxDataRow;
                for (int col = 0; col <= sheet.Cells.MaxDataColumn && dateColumnIndex == -1; col++)
                {
                    for (int row = 0; row <= maxRows; row++)
                    {
                        object val = sheet.Cells[row, col].Value;
                        if (val is DateTime)
                        {
                            dateColumnIndex = col;
                            break;
                        }
                    }
                }

                // If no date column is found, skip timeline addition for this file
                if (dateColumnIndex == -1)
                {
                    Console.WriteLine($"No Date column found in '{Path.GetFileName(sourcePath)}'. Skipping timeline.");
                    continue;
                }

                // ------------------------------------------------------------
                // 2. Create a PivotTable covering the used range of the worksheet
                // ------------------------------------------------------------
                int firstRow = sheet.Cells.MinDataRow;
                int firstCol = sheet.Cells.MinDataColumn;
                int lastRow = sheet.Cells.MaxDataRow;
                int lastCol = sheet.Cells.MaxDataColumn;

                string sourceRange = $"${CellIndexToName(firstRow, firstCol)}:${CellIndexToName(lastRow, lastCol)}";
                // Destination for the pivot table (place it a few rows below the data)
                string pivotDest = $"${CellIndexToName(lastRow + 2, 0)}";

                int pivotIndex = sheet.PivotTables.Add(sourceRange, pivotDest, "PivotTable_Timeline");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the date field to the Row area (base field for timeline)
                string dateFieldName = sheet.Cells[0, dateColumnIndex].StringValue; // assume header in first row
                if (string.IsNullOrEmpty(dateFieldName))
                    dateFieldName = $"Column{dateColumnIndex + 1}";

                pivot.AddFieldToArea(PivotFieldType.Row, dateFieldName);

                // Add another field (first non‑date column) to the Data area for demonstration
                for (int col = 0; col <= lastCol; col++)
                {
                    if (col == dateColumnIndex) continue;
                    string header = sheet.Cells[0, col].StringValue;
                    if (!string.IsNullOrEmpty(header))
                    {
                        pivot.AddFieldToArea(PivotFieldType.Data, header);
                        break;
                    }
                }

                // Refresh the pivot table so that it contains data
                pivot.RefreshData();
                pivot.CalculateData();

                // ------------------------------------------------------------
                // 3. Add a Timeline control linked to the pivot table
                // ------------------------------------------------------------
                // Place the timeline at cell E1 (row 0, column 4) using the date field name
                sheet.Timelines.Add(pivot, "E1", dateFieldName);

                // ------------------------------------------------------------
                // 4. Save the modified workbook to a temporary file
                // ------------------------------------------------------------
                string tempPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(sourcePath) + "_temp.xlsx");
                workbook.Save(tempPath, SaveFormat.Xlsx);

                // ------------------------------------------------------------
                // 5. Convert the temporary workbook to PDF using the provided ConversionUtility rule
                // ------------------------------------------------------------
                string pdfPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(sourcePath) + ".pdf");
                ConversionUtility.Convert(tempPath, pdfPath);

                // Clean up the temporary file
                File.Delete(tempPath);

                Console.WriteLine($"Processed '{Path.GetFileName(sourcePath)}' -> '{Path.GetFileName(pdfPath)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{Path.GetFileName(sourcePath)}': {ex.Message}");
            }
        }
    }

    // Helper: converts zero‑based row/column indices to Excel cell name (e.g., 0,0 => A1)
    private static string CellIndexToName(int row, int column)
    {
        // Column letters
        string colName = "";
        int dividend = column + 1;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            colName = Convert.ToChar('A' + modulo) + colName;
            dividend = (dividend - modulo) / 26;
        }
        // Row number (1‑based)
        return $"{colName}{row + 1}";
    }
}

public class Program
{
    // Application entry point
    public static void Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <inputDirectory> <outputDirectory>");
                return;
            }

            BatchExcelToPdfWithTimeline.Run(args[0], args[1]);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
