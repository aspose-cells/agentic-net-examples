// Title: C# – Split an Excel workbook into multiple CSV files by column groups using Aspose.Cells
// Description: Loads a source workbook, detects the last populated column, and iteratively copies a configurable number of columns (default 10) into new workbooks. Each new workbook is saved as a CSV file in a specified folder, producing a series of CSV parts that together represent the original sheet.
// Keywords: Aspose.Cells CSV split C# | Excel to CSV column groups | split worksheet by columns | C# copy columns Aspose.Cells | export Excel columns to separate CSV files | Aspose.Cells example GitHub | convert workbook to CSV programmatically
// Common Searches: How to split an Excel sheet into multiple CSV files by column count with Aspose.Cells | C# code to export selected column ranges to separate CSV files | Aspose.Cells copy columns and save as CSV example | GitHub Aspose.Cells CSV split sample | Split large Excel workbook into column‑wise CSV parts
// Developer Intent: Generate a series of CSV files, each containing a fixed number of columns taken from the original workbook.
// Use Cases: Create column‑wise CSV exports for downstream systems that require a maximum column width per file. | Distribute sections of a master workbook to different business units, each receiving only the columns they need. | Automate data‑migration pipelines where large sheets must be broken into smaller, column‑limited CSV chunks.
// AI Prompts: Modify the code to accept the column‑per‑file count as a command‑line argument. | Add a header row to every split CSV file while preserving column order. | Handle merged cells when copying column ranges for CSV export with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a source workbook, detects the last populated column, and iteratively copies a configurable number of columns (default 10) into new workbooks. Each new workbook is saved as a CSV file in a specified folder, producing a series of CSV parts that together represent the original sheet.
class Program
{
    static void Main()
    {
        // Path to the source workbook (can be any supported format)
        string sourcePath = "input.xlsx";

        // Folder where split CSV files will be saved
        string outputFolder = "SplitCsvOutput";
        Directory.CreateDirectory(outputFolder);

        // Load the source workbook
        Workbook sourceWorkbook = new Workbook(sourcePath);
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        Cells sourceCells = sourceSheet.Cells;

        // Determine the total number of columns that contain data
        int totalColumns = sourceCells.MaxDataColumn + 1; // zero‑based index + 1

        // Define how many columns each split CSV should contain
        int columnsPerFile = 10; // adjust as needed

        // Loop through the columns in groups
        int fileIndex = 1;
        for (int startCol = 0; startCol < totalColumns; startCol += columnsPerFile)
        {
            // Calculate how many columns to copy in this iteration
            int colsToCopy = Math.Min(columnsPerFile, totalColumns - startCol);

            // Create a new workbook for the current group
            Workbook splitWorkbook = new Workbook();
            Worksheet splitSheet = splitWorkbook.Worksheets[0];
            Cells splitCells = splitSheet.Cells;

            // Copy the selected column range from the source worksheet to the new workbook
            // Parameters: source cells, source start column, destination start column (0), number of columns
            sourceCells.CopyColumns(sourceCells, startCol, 0, colsToCopy);

            // Save the split workbook as CSV
            string csvFileName = Path.Combine(outputFolder, $"part_{fileIndex}.csv");
            splitWorkbook.Save(csvFileName, SaveFormat.Csv);

            fileIndex++;
        }

        // Clean up
        sourceWorkbook.Dispose();
    }
}
