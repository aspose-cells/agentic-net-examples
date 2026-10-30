// Title: Generate an Excel ListObject table from a CSV file with automatic column mapping using Aspose.Cells in C#
// Common Searches: how to load a CSV into a workbook and create a ListObject table with Aspose.Cells C# | Aspose.Cells C# add Excel table from CSV data range automatically | convert CSV to XLSX and define a table using ListObjects in Aspose.Cells | C# Aspose.Cells import CSV and set display name for ListObject | programmatically create Excel table from CSV using Aspose.Cells LoadOptions
// Tags: load CSV into Aspose.Cells workbook C# | add ListObject table programmatically | automatic column mapping for ListObject | CSV to XLSX conversion Aspose.Cells | set ListObject display name C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample checks for a CSV file, creates one if missing, loads it into an Aspose.Cells Workbook with LoadFormat.Csv, determines the used range, adds a ListObject (Excel table) over that range, assigns a display name, ensures the output directory exists, and saves the workbook as an XLSX file, all wrapped in exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for input CSV and output workbook.
            string csvPath = @"C:\Data\sample.csv";
            string outputPath = @"C:\Output\Result.xlsx";

            // Ensure the CSV file exists; create a simple one if missing.
            if (!File.Exists(csvPath))
            {
                string? csvDir = Path.GetDirectoryName(csvPath);
                if (!string.IsNullOrEmpty(csvDir))
                {
                    Directory.CreateDirectory(csvDir);
                }
                File.WriteAllText(csvPath, "Header1,Header2,Header3\nValue1,Value2,Value3");
            }

            // Load CSV into a workbook.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range after import.
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Define the area for the Excel table (ListObject).
            int startRow = 0;
            int startColumn = 0;
            int totalRows = maxRow - startRow + 1;
            int totalColumns = maxCol - startColumn + 1;

            // Add a ListObject (Excel table) over the imported range.
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            // Set a display name for the table (Name property not available in this version).
            table.DisplayName = "ImportedCsvTable";

            // Ensure the output directory exists.
            string? outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            // Save the workbook to an XLSX file.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
