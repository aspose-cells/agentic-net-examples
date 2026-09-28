// Title: Convert an Excel workbook to CSV and split the first worksheet into multiple files by column groups using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, partitions the first worksheet into fixed-size column blocks, and writes each block to a separate .csv file. | Create a reusable method in C# that accepts a source workbook path, destination folder, and a column‑per‑file count, then uses Aspose.Cells to copy the specified column ranges into new workbooks and save them as CSV. | Write a C# console program that iterates over columns of a worksheet, copies cell values and styles into temporary worksheets, and exports each segment as a CSV using Aspose.Cells SaveFormat.Csv.
// Common Searches: aspnet how to split an Excel sheet into multiple CSV files by column count using Aspose.Cells | c# Aspose.Cells export specific column range to separate CSV files | divide large worksheet into column groups and save each group as CSV with .NET | batch convert Excel columns to individual CSV files programmatically | Aspose.Cells save worksheet subset as CSV in C# console app
// Tags: Aspose.Cells split worksheet columns to CSV | C# export column groups as separate CSV files | Excel to CSV column-wise partitioning with Aspose | batch CSV generation from Excel columns .NET | copy cell styles when exporting to CSV Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookCsvSplitter
{
    // The program loads an Excel workbook, iterates over the first worksheet in configurable column groups, copies each group's cell values and styles into a new workbook, and saves each group as an individual CSV file in the specified output folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputFile = @"C:\Input\SourceWorkbook.xlsx";

            // Folder where split CSV files will be saved
            string outputFolder = @"C:\Output\CsvSplits";

            // Number of columns per split group
            int columnsPerGroup = 5;

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputFile);

            // Work with the first worksheet (adjust index if needed)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Determine total columns in the source sheet
            int totalColumns = sourceSheet.Cells.MaxColumn + 1;
            int totalRows = sourceSheet.Cells.MaxRow + 1;

            // Iterate over column groups
            for (int startCol = 0; startCol < totalColumns; startCol += columnsPerGroup)
            {
                // Calculate the end column for this group
                int endCol = Math.Min(startCol + columnsPerGroup - 1, totalColumns - 1);
                int groupColumnCount = endCol - startCol + 1;

                // Create a new workbook for the current group
                Workbook groupWorkbook = new Workbook();

                // Remove the default empty sheet and add a new one
                groupWorkbook.Worksheets.Clear();
                Worksheet groupSheet = groupWorkbook.Worksheets.Add("Sheet1");

                // Copy cells from the source sheet to the group sheet
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        // Source cell
                        Cell srcCell = sourceSheet.Cells[row, col];

                        // Destination cell (column index shifted to start at 0)
                        Cell destCell = groupSheet.Cells[row, col - startCol];

                        // Copy value
                        destCell.PutValue(srcCell.Value);

                        // Copy style (optional, but keeps formatting)
                        destCell.SetStyle(srcCell.GetStyle());
                    }
                }

                // Build output file name, e.g., "Part_1_5.csv" for columns 1-5
                string fileName = $"Part_{startCol + 1}_{endCol + 1}.csv";
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the group workbook as CSV
                groupWorkbook.Save(outputPath, SaveFormat.Csv);
            }

            Console.WriteLine("Workbook has been split and saved as CSV files.");
        }
    }
}
