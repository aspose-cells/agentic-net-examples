// Title: Import CSV into an Excel Table with Auto‑Detected Data Types using Aspose.Cells for .NET
// Description: C# example that validates a CSV file, creates a new Workbook, imports the CSV starting at A1 with automatic numeric conversion, determines the used range, builds a ListObject (Excel table) with headers, and saves the result as an XLSX file.
// Keywords: Aspose.Cells ImportCSV | CSV to Excel conversion .NET | automatic column type detection | create ListObject from CSV | Excel table generation C# | save workbook as XLSX | Aspose.Cells data import | autoConvert parameter | C# Excel automation | Aspose.Cells table example
// Common Searches: Aspose.Cells import CSV with auto data type detection | How to create an Excel table from a CSV using Aspose.Cells | C# ImportCSV autoConvert true example | Convert CSV to XLSX and add ListObject Aspose.Cells | Detect numeric columns when importing CSV in .NET
// Developer Intent: Read a CSV file, let Aspose.Cells infer column data types, wrap the imported range in an Excel table, and export the workbook as an XLSX file.
// Use Cases: Transform daily sales CSV files into structured Excel tables for pivot‑table analysis. | Build a configuration‑report generator that ingests CSV settings, preserves numeric formats, and outputs a formatted workbook. | Create a reusable utility that accepts any CSV, automatically types columns, adds a table with headers, and saves it for downstream processing.
// AI Prompts: Generate C# code with Aspose.Cells to import a CSV, enable autoConvert for data types, create a ListObject covering the data, and save as XLSX. | Explain how the ImportCSV method's autoConvert flag determines column types and how to retrieve the used range for table creation. | Suggest best‑practice error handling when loading a CSV and building an Excel table with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// C# example that validates a CSV file, creates a new Workbook, imports the CSV starting at A1 with automatic numeric conversion, determines the used range, builds a ListObject (Excel table) with headers, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the CSV file to be imported
            string csvPath = "data.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import CSV data starting at cell A1 (row 0, column 0)
            // Use comma as delimiter and enable automatic conversion of numeric strings
            cells.ImportCSV(csvPath, ",", true, 0, 0); // ImportCSV rule

            // Determine the used range after import
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
            int totalColumns = cells.MaxDataColumn + 1;

            // Add an Excel table (ListObject) over the imported range
            // hasHeaders = true assumes the first row contains column names
            int tableIndex = worksheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "ImportedCsvTable";

            // Save the workbook (lifecycle save rule)
            workbook.Save("ImportedTable.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as ImportedTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
