// Title: Convert an Excel workbook to separate CSV files per worksheet and generate a numeric summary report with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export each worksheet of an .xlsx to its own .csv file | How to create a summary text file with row, column and numeric totals using Aspose.Cells in .NET | C# code to save Excel worksheets as CSV and write workbook statistics to a text file | Generate per‑sheet CSV output and summary statistics from an Excel workbook with Aspose.Cells | Aspose.Cells example for iterating worksheets, saving as CSV and calculating numeric cell aggregates
// Tags: Aspose.Cells export worksheet to CSV | Aspose.Cells calculate numeric cell aggregates | Aspose.Cells generate workbook summary report | C# save individual worksheets as CSV files | Aspose.Cells used range row column count | Aspose.Cells create temporary workbook for single sheet | Aspose.Cells write summary statistics to text file

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, iterates through its worksheets, creates a temporary workbook for each sheet, and saves each one as an individual CSV file. It then computes total used rows and columns, counts numeric cells, sums their values, calculates the average, and writes these statistics to a summary text file placed alongside the generated CSV files.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Path to the source Excel workbook
            string sourcePath = @"C:\Data\SourceWorkbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Determine directory and base file name for output files
            string sourceDir = Path.GetDirectoryName(sourcePath) ?? Directory.GetCurrentDirectory();
            string baseFileName = Path.GetFileNameWithoutExtension(sourcePath);

            // Convert each worksheet to a separate CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    string csvPath = Path.Combine(sourceDir, $"{baseFileName}_{sheet.Name}.csv");

                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.Clear();

                    // Add a new sheet with the same name and copy contents
                    Worksheet tempSheet = tempWb.Worksheets.Add(sheet.Name);
                    sheet.Copy(tempSheet);

                    // Save the temporary workbook as CSV
                    tempWb.Save(csvPath, SaveFormat.Csv);
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{workbook.Worksheets[i].Name}': {exSheet.Message}");
                }
            }

            // Generate summary statistics
            int totalWorksheets = workbook.Worksheets.Count;
            long totalRows = 0;
            long totalColumns = 0;
            long numericCellCount = 0;
            double numericSum = 0.0;

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                totalRows += usedRange.RowCount;
                totalColumns += usedRange.ColumnCount;

                foreach (Cell cell in usedRange)
                {
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        numericCellCount++;
                        numericSum += cell.DoubleValue;
                    }
                }
            }

            double numericAverage = numericCellCount > 0 ? numericSum / numericCellCount : 0.0;

            // Build summary content
            string summary = $"Workbook Summary Statistics{Environment.NewLine}" +
                             $"--------------------------------{Environment.NewLine}" +
                             $"Source File: {Path.GetFileName(sourcePath)}{Environment.NewLine}" +
                             $"Total Worksheets: {totalWorksheets}{Environment.NewLine}" +
                             $"Total Rows (used): {totalRows}{Environment.NewLine}" +
                             $"Total Columns (used): {totalColumns}{Environment.NewLine}" +
                             $"Numeric Cells: {numericCellCount}{Environment.NewLine}" +
                             $"Sum of Numeric Values: {numericSum}{Environment.NewLine}" +
                             $"Average of Numeric Values: {numericAverage}{Environment.NewLine}";

            // Write summary to a text file alongside the CSV files
            string summaryPath = Path.Combine(sourceDir, $"{baseFileName}_Summary.txt");
            try
            {
                File.WriteAllText(summaryPath, summary);
            }
            catch (Exception exWrite)
            {
                Console.WriteLine($"Failed to write summary file: {exWrite.Message}");
            }

            Console.WriteLine("Conversion to CSV completed.");
            Console.WriteLine($"Summary statistics saved to: {summaryPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
