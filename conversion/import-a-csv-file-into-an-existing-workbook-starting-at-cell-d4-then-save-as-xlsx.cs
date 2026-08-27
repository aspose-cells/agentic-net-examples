// Title: Import a CSV file into an existing Excel workbook at cell D4 and save the result as XLSX using Aspose.Cells for .NET
// AI Prompts: Create C# code that opens an existing .xlsx workbook, uses Aspose.Cells Cells.ImportCSV to load a CSV file starting at row 4 column 4 with numeric conversion, and then saves the workbook as a new XLSX file. | Write a .NET example that checks for a CSV file, imports it into the first worksheet at cell D4 using a comma delimiter and the ImportCSV method, and exports the updated workbook to XLSX format with Aspose.Cells.
// Common Searches: Aspose.Cells C# import CSV into existing workbook starting at D4 | How to load a CSV file into a specific cell range with Aspose.Cells .NET | Save workbook as XLSX after importing CSV data using Aspose.Cells | Import CSV with numeric conversion into Excel using Aspose.Cells C# | Using Cells.ImportCSV to place data at cell D4 in Aspose.Cells
// Tags: cells.importcsv method start cell d4 | load existing xlsx workbook aspose.cells | export workbook to xlsx aspose.cells | csv numeric conversion aspose.cells | import csv data into worksheet aspose.cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImportDemo
{
    // The example loads an existing Excel file (or creates a new one), imports a CSV file into the first worksheet beginning at cell D4 with numeric conversion using a comma delimiter, and saves the modified workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the existing workbook that will receive the CSV data
                string existingWorkbookPath = "ExistingWorkbook.xlsx";

                Workbook workbook;

                // Load the existing workbook if it exists; otherwise create a new one
                if (File.Exists(existingWorkbookPath))
                {
                    workbook = new Workbook(existingWorkbookPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Ensure at least one worksheet exists
                    workbook.Worksheets.Add();
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Path to the CSV file to be imported
                string csvFilePath = "DataFile.csv";

                // Verify CSV file existence before importing
                if (File.Exists(csvFilePath))
                {
                    // Import CSV data starting at cell D4 (row index 3, column index 3)
                    // Using comma as the delimiter and converting numeric strings to numbers
                    cells.ImportCSV(csvFilePath, ",", true, 3, 3);
                }
                else
                {
                    Console.WriteLine($"CSV file not found: '{csvFilePath}'. Skipping import.");
                }

                // Save the modified workbook as XLSX
                string outputPath = "WorkbookWithCsv.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
