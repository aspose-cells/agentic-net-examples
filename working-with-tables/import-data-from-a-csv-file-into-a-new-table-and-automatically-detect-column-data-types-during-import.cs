// Title: Import a CSV file into an Excel workbook as a ListObject table with automatic column type detection using Aspose.Cells for .NET
// AI Prompts: Load a CSV file into a Workbook with LoadOptions(LoadFormat.Csv) so Aspose.Cells infers column data types, then obtain the worksheet's used range. | Create a ListObject on the first worksheet covering the used range, mark the first row as headers, assign a display name, and save the workbook as an .xlsx file.
// Common Searches: how to import csv into excel as a table with Aspose.Cells C# | Aspose.Cells automatically detect column types when loading CSV | create ListObject from CSV data in Aspose.Cells .NET | save imported CSV as Excel table with headers using Aspose.Cells | C# load CSV to workbook and convert to Excel table with type inference
// Tags: load csv with automatic type detection Aspose.Cells | create ListObject table from worksheet range C# | save workbook as xlsx after csv import Aspose.Cells | detect column data types during csv load Aspose.Cells | add table with headers from csv using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example checks for the CSV file, loads it into an Aspose.Cells Workbook using LoadOptions so column data types are auto‑detected, creates a ListObject that spans the used range with the first row as headers, names the table "ImportedTable", and saves the result as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source CSV file
            string csvPath = "data.csv";

            // Ensure the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: The file '{csvPath}' was not found.");
                return;
            }

            // Load the CSV file into a new workbook.
            // Aspose.Cells automatically detects column data types while loading CSV.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // The CSV data is placed in the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet.
            int firstRow = sheet.Cells.MinRow;
            int firstColumn = sheet.Cells.MinColumn;
            int lastRow = sheet.Cells.MaxDataRow;
            int lastColumn = sheet.Cells.MaxDataColumn;

            // Create a table (ListObject) that covers the used range.
            // The 'true' argument indicates that the first row contains column headers.
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "ImportedTable";

            // Save the workbook with the new table to an Excel file.
            string outputPath = "ImportedTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
