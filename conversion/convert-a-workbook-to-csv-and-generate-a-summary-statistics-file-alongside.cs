// Title: Convert an Excel workbook to individual UTF-8 CSV files per worksheet and create a summary statistics text file with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, iterates through all worksheets, and saves each one as a UTF-8 encoded CSV file. | Add logic to compute total worksheets, used rows, used columns, and total used cells, then write these metrics to a text summary file in the same output directory.
// Common Searches: Aspose.Cells C# export each sheet of an Excel file to separate CSV files | How to generate a workbook summary (total rows, columns, cells) with Aspose.Cells .NET | Save Excel worksheets as UTF-8 CSV using TxtSaveOptions in C# | Create a text report of used range statistics after converting Excel to CSV with Aspose.Cells | Batch convert multi-sheet workbook to CSV and get summary statistics in .NET
// Tags: Aspose.Cells export worksheet to CSV | TxtSaveOptions UTF8 CSV Aspose.Cells | calculate used range rows columns cells Aspose.Cells | workbook summary statistics text file .NET | batch convert Excel workbook to multiple CSV files

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an .xlsx workbook with Aspose.Cells, saves each worksheet as a UTF-8 CSV file using TxtSaveOptions, computes total worksheets, used rows, columns, and cells across all sheets, and writes these metrics to a WorkbookSummary.txt file in the output folder.
class WorkbookToCsvWithSummary
{
    static void Main(string[] args)
    {
        // Input Excel file path
        string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        // Output folder for CSV files and summary
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare summary data
            int worksheetCount = workbook.Worksheets.Count;
            int totalRows = 0;
            int totalColumns = 0;
            long totalCells = 0;

            // Iterate through each worksheet
            for (int i = 0; i < worksheetCount; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                string sheetName = sheet.Name;

                // Determine the used range (fully qualified to avoid ambiguity)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int rows = usedRange?.RowCount ?? 0;
                int columns = usedRange?.ColumnCount ?? 0;

                totalRows += rows;
                totalColumns += columns;
                totalCells += (long)rows * columns;

                // Define CSV save options
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Encoding = Encoding.UTF8,
                    Separator = ','
                };

                // Build CSV file path (sanitize file name)
                string safeSheetName = string.Join("_", sheetName.Split(Path.GetInvalidFileNameChars()));
                string csvPath = Path.Combine(outputFolder, $"{safeSheetName}.csv");

                // Save only the current worksheet as CSV by setting it as active
                workbook.Worksheets.ActiveSheetIndex = i;
                workbook.Save(csvPath, csvOptions);
            }

            // Build summary content
            string summary = $"Workbook Summary Statistics{Environment.NewLine}" +
                             $"--------------------------------{Environment.NewLine}" +
                             $"Total Worksheets: {worksheetCount}{Environment.NewLine}" +
                             $"Total Used Rows (across all sheets): {totalRows}{Environment.NewLine}" +
                             $"Total Used Columns (across all sheets): {totalColumns}{Environment.NewLine}" +
                             $"Total Cells (used): {totalCells}{Environment.NewLine}" +
                             $"Generated on: {DateTime.Now}{Environment.NewLine}";

            // Write summary to a text file alongside CSV files
            string summaryPath = Path.Combine(outputFolder, "WorkbookSummary.txt");
            File.WriteAllText(summaryPath, summary);

            Console.WriteLine("Conversion to CSV completed. Summary file created at:");
            Console.WriteLine(summaryPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
