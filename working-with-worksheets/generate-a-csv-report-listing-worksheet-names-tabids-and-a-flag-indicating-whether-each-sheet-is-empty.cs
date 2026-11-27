// Title: Generate a CSV file with worksheet names, TabId values, and empty‑sheet flag using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export worksheet name, TabId and empty status to CSV | How to determine if an Excel worksheet is empty with Aspose.Cells and write the result to a CSV file | Create a worksheet metadata report (name, TabId, IsEmpty) using Aspose.Cells in a .NET console application | Iterate through worksheets with Aspose.Cells and output TabId and emptiness flag to CSV | C# console program to generate CSV report of Excel sheets including empty‑sheet detection
// Tags: Aspose.Cells worksheet TabId extraction | empty worksheet detection Aspose.Cells | CSV generation from Excel worksheets C# | iterate worksheets Aspose.Cells .NET | export worksheet metadata to CSV

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetReportGenerator
{
    // The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, retrieves the sheet name and TabId, checks if the sheet is empty by verifying MaxDataRow and MaxDataColumn are negative, escapes commas in names, and writes these details to a UTF‑8 CSV file with a header row.
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length < 2)
            {
                Console.WriteLine("Usage: WorksheetReportGenerator <inputWorkbookPath> <outputCsvPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at \"{Path.GetFullPath(inputPath)}\".");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputPath);

                // Prepare the CSV file for writing
                using (StreamWriter writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8))
                {
                    // Write CSV header
                    writer.WriteLine("WorksheetName,TabId,IsEmpty");

                    // Iterate through each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        string sheetName = sheet.Name;
                        int tabId = sheet.TabId;

                        // Determine if the sheet is empty.
                        // A sheet is considered empty when there are no cells with data.
                        // MaxDataRow and MaxDataColumn return -1 when no data exists.
                        bool isEmpty = sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0;

                        // Escape commas in sheet names if any
                        if (sheetName.Contains(","))
                        {
                            sheetName = $"\"{sheetName}\"";
                        }

                        // Write the CSV line
                        writer.WriteLine($"{sheetName},{tabId},{isEmpty}");
                    }
                }

                Console.WriteLine($"CSV report generated successfully at \"{Path.GetFullPath(outputPath)}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
