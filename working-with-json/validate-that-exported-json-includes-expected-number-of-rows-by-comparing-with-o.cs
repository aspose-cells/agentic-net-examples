using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonRowValidation
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and populate data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some data with a few empty rows in between
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            // Row 3 left empty
            sheet.Cells["A4"].PutValue(2);
            sheet.Cells["B4"].PutValue("Bob");
            sheet.Cells["A5"].PutValue(3);
            sheet.Cells["B5"].PutValue("Charlie");

            // Original row count (including empty rows up to the last used row)
            int originalRowCount = sheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based
            Console.WriteLine($"Original worksheet row count: {originalRowCount}");

            // ---------- Export the worksheet to JSON ----------
            string jsonPath = "workbook.json";
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                SkipEmptyRows = false,          // Ensure empty rows are kept in the JSON
                ExportNestedStructure = true,   // Export as nested objects (default)
                HasHeaderRow = true
            };
            workbook.Save(jsonPath, saveOptions);
            Console.WriteLine($"Workbook exported to JSON at: {Path.GetFullPath(jsonPath)}");

            // ---------- Load the JSON back into a new workbook ----------
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                MultipleWorksheets = true       // Preserve each sheet as a separate worksheet
            };
            Workbook jsonWorkbook = new Workbook(jsonPath, loadOptions);
            Worksheet jsonSheet = jsonWorkbook.Worksheets[0];

            // Row count after loading from JSON
            int jsonRowCount = jsonSheet.Cells.MaxDataRow + 1;
            Console.WriteLine($"Row count after importing JSON: {jsonRowCount}");

            // ---------- Validation ----------
            if (originalRowCount == jsonRowCount)
            {
                Console.WriteLine("Validation succeeded: Exported JSON contains the expected number of rows.");
            }
            else
            {
                Console.WriteLine($"Validation failed: Expected {originalRowCount} rows, but JSON contains {jsonRowCount} rows.");
            }
        }
    }
}