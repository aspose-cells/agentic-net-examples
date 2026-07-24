// Title: Split an Excel workbook into multiple CSV files by column count with Aspose.Cells for .NET
// Description: Loads an Excel workbook, groups its columns into configurable batches, copies each batch to a new workbook, and saves each as a separate CSV file in a specified folder using Aspose.Cells for C#.
// Keywords: Aspose.Cells split columns CSV | C# export Excel columns to CSV | convert workbook to multiple CSV files | copy column range Aspose.Cells | save worksheet as CSV .NET | Excel to CSV batch export | column‑wise CSV split C#
// Common Searches: how to split Excel sheet into several CSV files by column count using Aspose.Cells | C# code to export specific column groups from a workbook to separate CSV files | Aspose.Cells copy selected columns and save as CSV | divide Excel columns into multiple CSV files programmatically | Aspose.Cells CSV split example .NET
// Developer Intent: Generate separate CSV files for each defined set of columns in an Excel workbook using Aspose.Cells.
// Use Cases: Create department‑specific CSV reports when each department's data occupies a fixed number of columns in a master sheet. | Break large spreadsheets into column‑limited chunks for systems that impose a maximum column width. | Automate CSV slicing for visualization tools that require one file per data segment.
// AI Prompts: Write C# code with Aspose.Cells that splits a worksheet into CSV files, each containing a configurable number of columns. | Explain how to copy a range of columns from one worksheet to a new workbook and export it as CSV using Aspose.Cells. | Add robust error handling for missing source files, empty worksheets, and invalid column counts when splitting an Excel workbook into multiple CSV files.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookToCsvSplit
{
    // Loads an Excel workbook, groups its columns into configurable batches, copies each batch to a new workbook, and saves each as a separate CSV file in a specified folder using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Folder where split CSV files will be saved
            string outputFolder = "SplitCsvOutput";
            Directory.CreateDirectory(outputFolder);

            // Number of columns per split CSV file
            int columnsPerFile = 5; // adjust as needed

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Determine total number of data columns in the source sheet
            int totalColumns = sourceSheet.Cells.MaxDataColumn + 1;

            // Iterate over column groups and create separate CSV files
            for (int startCol = 0; startCol < totalColumns; startCol += columnsPerFile)
            {
                // Calculate how many columns to copy for this group
                int colsToCopy = Math.Min(columnsPerFile, totalColumns - startCol);

                // Create a new workbook for the current group
                Workbook splitWorkbook = new Workbook();
                Worksheet splitSheet = splitWorkbook.Worksheets[0];

                // Copy the selected column range from the source sheet to the new workbook
                // Parameters: source cells, source start column, destination start column, number of columns
                splitSheet.Cells.CopyColumns(sourceSheet.Cells, startCol, 0, colsToCopy);

                // Build the output CSV file name
                string csvFileName = $"Part_{(startCol / columnsPerFile) + 1}.csv";
                string csvPath = Path.Combine(outputFolder, csvFileName);

                // Save the split workbook as CSV
                splitWorkbook.Save(csvPath, SaveFormat.Csv);
            }

            Console.WriteLine("Workbook has been split into CSV files successfully.");
        }
    }
}
