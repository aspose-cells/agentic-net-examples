// Title: C# – Import CSV and Create an Excel ListObject (Query Table) with Automatic Column Mapping using Aspose.Cells
// Description: Shows how to read an external CSV file with Aspose.Cells for .NET, automatically convert data types, determine the used range, add a ListObject that uses the first row as headers, give the table a display name, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# | ImportCSV | ListObject | Query Table | CSV to Excel | automatic column mapping | Excel table creation | XLSX export | Workbook automation
// Common Searches: Aspose.Cells import CSV to ListObject C# | Create Excel table from CSV using Aspose.Cells | C# code for CSV to Excel query table | Automatic column mapping Aspose.Cells .NET | How to add a ListObject from CSV with Aspose
// Developer Intent: Programmatically generate an Excel table from a CSV file with headers and mapped columns.
// Use Cases: Transform daily sales CSV exports into structured Excel tables for pivot‑table reporting. | Convert configuration or lookup CSV files into filterable query tables within automated reports. | Batch‑process log CSVs, creating a ListObject for each to enable Excel‑based validation, charting, and analysis.
// AI Prompts: Write C# code that reads a CSV file, imports it with ImportCSV, creates a ListObject over the exact used range, sets the first row as headers, assigns a custom display name, and saves the workbook as XLSX. | Provide a reusable method `Workbook CreateWorkbookFromCsv(string csvPath, string tableName)` that returns a workbook containing a query table with automatic column mapping. | Explain how to detect the populated range after ImportCSV, add a ListObject that spans that range, and handle missing output directories gracefully.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsQueryTableFromCsv
{
    // Shows how to read an external CSV file with Aspose.Cells for .NET, automatically convert data types, determine the used range, add a ListObject that uses the first row as headers, give the table a display name, and save the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the external CSV file
                string csvPath = "data.csv";

                // Verify that the CSV file exists to avoid FileNotFoundException
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Import the CSV data starting at cell A1 (row 0, column 0)
                // Using comma as delimiter and converting numeric data automatically
                cells.ImportCSV(csvPath, ",", true, 0, 0);

                // Determine the range that now contains the imported data
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

                // Calculate start and end positions for the ListObject
                int startRow = usedRange.FirstRow;
                int startColumn = usedRange.FirstColumn;
                int endRow = startRow + usedRange.RowCount - 1;
                int endColumn = startColumn + usedRange.ColumnCount - 1;

                // Add a ListObject (Excel table) over the imported range
                // The last parameter 'true' indicates that the first row contains column headers
                int listObjectIndex = sheet.ListObjects.Add(
                    startRow, startColumn,
                    endRow, endColumn,
                    true);

                ListObject table = sheet.ListObjects[listObjectIndex];

                // Optionally give the table a display name
                table.DisplayName = "CsvDataTable";

                // Save the workbook to an XLSX file
                string outputPath = "CsvDataAsQueryTable.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
