// Title: Export a QueryTable from an Excel worksheet to a CSV file with headers using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, retrieves the first QueryTable on the first worksheet, and writes its result range—including column headers—to a CSV file with Aspose.Cells. | Show how to copy a QueryTable's result range into a new workbook and save it as CSV using TxtSaveOptions while preserving original data types in .NET.
// Common Searches: how to extract the first query table from an Excel file and save it as CSV with Aspose.Cells C# | preserving data types and headers when exporting a query table to CSV using Aspose.Cells | Aspose.Cells copy query table result range to new workbook before CSV export | C# Aspose.Cells TxtSaveOptions settings for CSV export of query tables | export querytable to csv without losing column formatting Aspose.Cells .NET
// Tags: export querytable to csv Aspose.Cells | copy querytable result range Aspose.Cells | TxtSaveOptions CSV export .NET | preserve column headers Aspose.Cells | querytable data type retention C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// Loads an Excel workbook, obtains the first QueryTable on the first worksheet, copies its result range (including headers) to a new workbook, and saves the data as a CSV file using TxtSaveOptions.
class ExportQueryTableToCsv
{
    static void Main()
    {
        try
        {
            // Input and output file paths.
            const string sourcePath = "SourceWithQueryTable.xlsx";
            const string outputPath = "ExportedQueryTable.csv";

            // Verify that the source workbook exists.
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the workbook that contains the query table.
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Assume the query table is on the first worksheet.
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Ensure there is at least one query table.
            if (sourceSheet.QueryTables.Count == 0)
                throw new InvalidOperationException("No query tables found on the first worksheet.");

            // Get the first query table.
            QueryTable queryTable = sourceSheet.QueryTables[0];

            // Get the result range of the query table (includes headers).
            AsposeRange resultRange = queryTable.ResultRange;

            // Calculate row and column counts.
            int rowCount = resultRange.RowCount;
            int columnCount = resultRange.ColumnCount;

            // Create a new workbook to hold only the query table data.
            Workbook exportWorkbook = new Workbook();
            exportWorkbook.Worksheets.Clear(); // Remove default sheet.
            Worksheet exportSheet = exportWorkbook.Worksheets.Add("Export");

            // Copy the query table range to the new workbook.
            AsposeRange srcRange = sourceSheet.Cells.CreateRange(
                resultRange.FirstRow, resultRange.FirstColumn, rowCount, columnCount);
            AsposeRange destRange = exportSheet.Cells.CreateRange(0, 0, rowCount, columnCount);
            destRange.Copy(srcRange);

            // Prepare CSV save options using TxtSaveOptions (compatible with all versions).
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Separator = ',', // Use comma as separator.
                // The range already contains headers, so no extra options needed.
            };

            // Save the export workbook as CSV.
            exportWorkbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Query table exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
