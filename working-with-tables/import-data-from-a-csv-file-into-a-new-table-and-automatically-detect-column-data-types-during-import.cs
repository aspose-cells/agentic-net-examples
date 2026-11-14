// Title: Import a CSV file into a new ListObject table with automatic column type detection using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load CSV file and convert to Excel table | how to automatically detect column data types when importing CSV with Aspose.Cells | create a ListObject in a worksheet from CSV data using Aspose.Cells .NET | convert CSV to XLSX and add a table with headers using Aspose.Cells C# | Aspose.Cells example for importing CSV into a new table with auto‑fit columns
// Tags: load csv into workbook Aspose.Cells | create listobject from csv data | auto detect column types csv Aspose.Cells | add excel table programmatically C# | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample verifies the CSV file exists, loads it into a Workbook with CSV LoadOptions, determines the used range, adds a ListObject covering that range with headers, assigns a display name, auto‑fits columns, and saves the result as an XLSX file, handling missing directories and exceptions.
class CsvToTableImporter
{
    static void Main()
    {
        try
        {
            // Path to the CSV file to import
            string csvPath = "input.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: CSV file not found at path '{csvPath}'.");
                return;
            }

            // Load the CSV file into a new workbook.
            // If a newer Aspose.Cells version provides CsvLoadOptions, it can be used here.
            // For compatibility, we use the generic LoadOptions with CSV format.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Get the first worksheet where the CSV data has been loaded
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range after import
            int lastRow = sheet.Cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = sheet.Cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Add a new table (ListObject) covering the imported range
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, lastRow + 1, lastColumn + 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table
            table.DisplayName = "ImportedDataTable";

            // Optional: Auto‑fit columns for better appearance
            sheet.AutoFitColumns();

            // Save the workbook to an Excel file
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
