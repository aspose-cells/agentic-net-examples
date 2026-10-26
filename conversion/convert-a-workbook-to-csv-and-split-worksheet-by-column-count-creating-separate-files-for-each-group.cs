// Title: Convert an Excel workbook to CSV and split each worksheet into column‑grouped files with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# split worksheet columns into separate CSV files by fixed column count | How to export parts of an Excel sheet as multiple CSVs using Aspose.Cells .NET | C# code to divide a workbook into CSV chunks based on number of columns per file | Save each column group of an Excel workbook as individual CSV files with Aspose.Cells | Create CSV partitions from Excel worksheets by column range in C#
// Tags: Aspose.Cells column group CSV generation | C# batch export of worksheet sections to CSV | Excel column range CSV partition with Aspose | retain cell style in CSV export Aspose.Cells | automated CSV splitting of workbook worksheets .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, determines the used rows and columns for each worksheet, and then divides the columns into groups of a configurable size (e.g., five columns per group). For each group a new workbook is created, cell values and styles are copied, and the workbook is saved as a CSV file in a specified output folder, producing separate CSV files for each column segment.
class WorkbookCsvSplitter
{
    static void Main()
    {
        // Input workbook path
        string inputPath = @"C:\Input\SourceWorkbook.xlsx";

        // Folder where split CSV files will be saved
        string outputFolder = @"C:\Output\CsvParts";

        // Number of columns per split group
        int columnsPerGroup = 5;

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Load the source workbook
        Workbook sourceWb = new Workbook(inputPath);

        // Iterate through each worksheet in the source workbook
        foreach (Worksheet sourceSheet in sourceWb.Worksheets)
        {
            // Determine the total number of columns that contain data
            int totalColumns = sourceSheet.Cells.MaxColumn + 1; // MaxColumn is zero‑based
            int totalRows = sourceSheet.Cells.MaxRow + 1;       // MaxRow is zero‑based

            // Process columns in groups defined by columnsPerGroup
            for (int startCol = 0; startCol < totalColumns; startCol += columnsPerGroup)
            {
                int endCol = Math.Min(startCol + columnsPerGroup - 1, totalColumns - 1);
                int groupIndex = (startCol / columnsPerGroup) + 1;

                // Create a new workbook for the current column group
                Workbook partWb = new Workbook();
                Worksheet partSheet = partWb.Worksheets[0];
                partSheet.Name = $"{sourceSheet.Name}_Part{groupIndex}";

                // Copy cells from the source sheet to the part sheet
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell sourceCell = sourceSheet.Cells[row, col];
                        Cell targetCell = partSheet.Cells[row, col - startCol];

                        // Copy value
                        targetCell.PutValue(sourceCell.Value);

                        // Copy style (optional but preserves formatting)
                        targetCell.SetStyle(sourceCell.GetStyle());
                    }
                }

                // Build output CSV file name
                string csvFileName = $"{sourceSheet.Name}_Part{groupIndex}.csv";
                string csvPath = Path.Combine(outputFolder, csvFileName);

                // Save the part workbook as CSV
                partWb.Save(csvPath, SaveFormat.Csv);
            }
        }

        Console.WriteLine("Workbook has been split and saved as CSV files.");
    }
}
