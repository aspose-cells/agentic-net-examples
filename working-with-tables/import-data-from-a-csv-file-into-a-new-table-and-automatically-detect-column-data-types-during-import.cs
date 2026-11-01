// Title: Import a CSV file into an Excel worksheet as a ListObject table with auto-detected column types using Aspose.Cells for .NET
// AI Prompts: Load a CSV file with LoadOptions(LoadFormat.Csv) into a Workbook, calculate the used range, and add a ListObject that automatically infers each column's data type. | Assign a display name to the ListObject, enable the header row and stripe styling, then save the workbook as an XLSX file.
// Common Searches: how to convert csv to excel table with Aspose.Cells .NET and keep data types | Aspose.Cells create ListObject from CSV file automatically detecting column types | apply built-in table style when importing CSV into Excel using Aspose.Cells
// Tags: Aspose.Cells CSV to ListObject conversion | auto detect column data types Aspose.Cells | apply TableStyleMedium9 ListObject | use LoadOptions for CSV import Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example checks for the existence of input.csv, loads it into a Workbook using LoadFormat.Csv, determines the used range, creates a ListObject covering that range with automatically detected column types, configures table properties and a built-in style, and finally saves the result as output.xlsx.
class CsvToExcelTable
{
    static void Main()
    {
        // Paths for input CSV and output Excel files
        string csvPath = "input.csv";
        string excelPath = "output.xlsx";

        // Verify that the CSV file exists to avoid FileNotFoundException
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
            return;
        }

        try
        {
            // Load the CSV file using the correct LoadFormat enum value
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // The CSV data is loaded into the first worksheet by default
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range that contains the imported data
            int firstRow = sheet.Cells.MinRow;
            int firstColumn = sheet.Cells.MinColumn;
            int totalRows = sheet.Cells.MaxRow - firstRow + 1;
            int totalColumns = sheet.Cells.MaxColumn - firstColumn + 1;

            // Add a new table (ListObject) covering the used range.
            // The Add method returns the index of the created table.
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set table properties
            table.DisplayName = "ImportedData";
            table.ShowHeaderRow = true;
            table.ShowTableStyleFirstColumn = false;
            table.ShowTableStyleLastColumn = false;
            table.ShowTableStyleRowStripes = true;
            table.ShowTableStyleColumnStripes = false;

            // Apply a built‑in table style for better visual appearance
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Successfully converted '{csvPath}' to '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
