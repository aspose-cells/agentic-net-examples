using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvSplitExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Load the workbook using the standard constructor (lifecycle rule)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define how many columns each split CSV should contain
            int columnsPerFile = 5; // adjust as needed

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
            {
                // Determine the total number of columns that contain data
                int totalColumns = sourceSheet.Cells.MaxColumn + 1; // MaxColumn is zero‑based

                // Calculate how many groups are needed
                int groupCount = (totalColumns + columnsPerFile - 1) / columnsPerFile;

                for (int groupIndex = 0; groupIndex < groupCount; groupIndex++)
                {
                    // Create a new workbook for the current column group (lifecycle rule)
                    Workbook partWorkbook = new Workbook();

                    // Get the default worksheet of the new workbook
                    Worksheet partSheet = partWorkbook.Worksheets[0];

                    // Calculate the start column and how many columns to copy for this group
                    int startColumn = groupIndex * columnsPerFile;
                    int columnsToCopy = Math.Min(columnsPerFile, totalColumns - startColumn);

                    // Copy the selected columns from the source sheet to the part sheet
                    // Using Cells.CopyColumns(source, sourceColumn, destinationColumn, columnCount)
                    partSheet.Cells.CopyColumns(sourceSheet.Cells, startColumn, 0, columnsToCopy);

                    // Build the output CSV file name
                    string outputFileName = $"{Path.GetFileNameWithoutExtension(sourcePath)}_{sourceSheet.Name}_Part{groupIndex + 1}.csv";

                    // Save the part workbook as CSV (save rule)
                    partWorkbook.Save(outputFileName, SaveFormat.Csv);

                    // Dispose the part workbook to free resources
                    partWorkbook.Dispose();
                }
            }

            // Dispose the source workbook
            sourceWorkbook.Dispose();

            Console.WriteLine("Workbook has been split into CSV files successfully.");
        }
    }
}