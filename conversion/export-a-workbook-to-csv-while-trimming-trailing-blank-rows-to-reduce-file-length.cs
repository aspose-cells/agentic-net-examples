// Title: Export an Excel workbook to CSV with Aspose.Cells while removing trailing blank rows in C#
// Common Searches: Aspose.Cells C# export worksheet to CSV after deleting empty rows at the bottom | how to trim trailing blank rows in Excel file before saving as CSV using Aspose.Cells | remove extra rows beyond MaxDataRow and save as UTF‑8 CSV with Aspose.Cells | C# code to delete rows after last data row in each sheet and convert to CSV | save trimmed Excel workbook as CSV with UTF-8 encoding using Aspose.Cells
// Tags: delete rows after MaxDataRow Aspose.Cells | export trimmed workbook to CSV UTF-8 | Aspose.Cells CSV conversion with row trimming | remove trailing blank rows Excel C# | TxtSaveOptions CSV encoding UTF8 Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an existing XLSX file, iterates through each worksheet to find the last populated row (MaxDataRow), deletes any rows below it, and then saves the cleaned workbook as a UTF‑8 encoded CSV using TxtSaveOptions.
class ExportWorkbookToCsv
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Trim trailing blank rows in each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last row that contains data (zero‑based index)
                int lastDataRow = sheet.Cells.MaxDataRow;

                // Index of the last possible row in the sheet (zero‑based)
                int totalRows = sheet.Cells.MaxRow;

                // If there are rows after the last data row, delete them
                if (lastDataRow < totalRows)
                {
                    int rowsToDelete = totalRows - lastDataRow;
                    sheet.Cells.DeleteRows(lastDataRow + 1, rowsToDelete);
                }
            }

            // Save the trimmed workbook as CSV with UTF‑8 encoding
            string outputPath = "output.csv";
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Encoding = Encoding.UTF8
            };
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Workbook successfully exported to CSV: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
