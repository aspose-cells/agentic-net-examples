// Title: Split an Excel workbook into multiple CSV files by column groups with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# split worksheet columns into separate CSV files | How to export Excel columns in groups to CSV using Aspose.Cells .NET | C# code to divide a workbook into CSV parts based on column count | Save each column block of an Excel sheet as individual CSV with Aspose.Cells | Batch export column ranges to CSV files from an Excel workbook in C#
// Tags: Aspose.Cells export column groups to CSV | C# split worksheet columns into multiple CSV files | Excel to CSV conversion with column batching using Aspose | Workbook column range CSV split .NET | Save part of worksheet as CSV Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates each worksheet, and partitions its columns into groups of a specified size. For each group it exports the data to a DataTable, copies the values into a new workbook, and saves the result as a CSV file named with the worksheet and part index.
class WorkbookCsvSplitter
{
    // Splits each worksheet of the source workbook into CSV files,
    // each containing a group of columns defined by columnGroupSize.
    public static void SplitWorkbookToCsv(string sourceFilePath, string outputDirectory, int columnGroupSize)
    {
        try
        {
            // Verify source file exists
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException($"Source file not found: {sourceFilePath}");

            // Ensure output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourceFilePath);

            // Iterate through all worksheets in the source workbook
            foreach (Worksheet srcWorksheet in sourceWorkbook.Worksheets)
            {
                // Determine the total number of rows and columns with data
                int totalRows = srcWorksheet.Cells.MaxDataRow + 1;      // zero‑based index + 1
                int totalColumns = srcWorksheet.Cells.MaxColumn + 1;   // zero‑based index + 1

                // Process columns in groups of columnGroupSize
                int groupIndex = 1;
                for (int startCol = 0; startCol < totalColumns; startCol += columnGroupSize, groupIndex++)
                {
                    // Number of columns to copy for this group
                    int colsToCopy = Math.Min(columnGroupSize, totalColumns - startCol);

                    // Export the selected column range to a DataTable
                    DataTable dt = srcWorksheet.Cells.ExportDataTable(
                        0,                     // start row
                        startCol,              // start column
                        totalRows,             // total rows
                        colsToCopy,            // total columns
                        true);                 // export column names

                    // Create a new workbook for the CSV output
                    Workbook destWorkbook = new Workbook();
                    Worksheet destWorksheet = destWorkbook.Worksheets[0];

                    // Manually copy DataTable values into the destination worksheet
                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            destWorksheet.Cells[r, c].Value = dt.Rows[r][c];
                        }
                    }

                    // Build the output CSV file name
                    string csvFileName = $"{srcWorksheet.Name}_Part{groupIndex}.csv";
                    string csvFilePath = Path.Combine(outputDirectory, csvFileName);

                    // Save the workbook as CSV
                    destWorkbook.Save(csvFilePath, SaveFormat.Csv);
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during workbook split: {ex.Message}");
        }
    }

    // Example usage
    static void Main()
    {
        try
        {
            string sourcePath = @"C:\Input\SampleWorkbook.xlsx";
            string outputDir = @"C:\Output\CsvParts";
            int columnsPerFile = 5; // Adjust the column group size as needed

            SplitWorkbookToCsv(sourcePath, outputDir, columnsPerFile);
            Console.WriteLine("Workbook has been split into CSV files.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
