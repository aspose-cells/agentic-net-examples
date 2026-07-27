using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonRowValidation
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and populate sample data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");

            // Data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");

            // Intentionally leave row 4 empty
            // Add another data row after the empty row
            cells["A5"].PutValue(3);
            cells["B5"].PutValue("Charlie");

            // Original row count (including empty rows up to the last used row)
            int originalRowCount = cells.MaxDataRow + 1; // MaxDataRow is zero‑based

            // ---------- Export the worksheet to JSON ----------
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Do not skip empty rows so the row count in JSON matches the worksheet
                SkipEmptyRows = false,
                // Export as a single JSON object (optional, but keeps structure simple)
                AlwaysExportAsJsonObject = true
            };

            string jsonPath = "workbook.json";
            workbook.Save(jsonPath, saveOptions);

            // ---------- Load the JSON back into a new workbook ----------
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                // Load each JSON attribute as a separate worksheet if needed (default false)
                MultipleWorksheets = false
            };

            Workbook jsonWorkbook = new Workbook(jsonPath, loadOptions);
            Worksheet jsonSheet = jsonWorkbook.Worksheets[0];
            Cells jsonCells = jsonSheet.Cells;

            // Row count after loading JSON
            int jsonRowCount = jsonCells.MaxDataRow + 1;

            // ---------- Validate row counts ----------
            Console.WriteLine($"Original worksheet row count : {originalRowCount}");
            Console.WriteLine($"Row count after JSON export/load : {jsonRowCount}");

            if (originalRowCount == jsonRowCount)
            {
                Console.WriteLine("Validation succeeded: Row counts match.");
            }
            else
            {
                Console.WriteLine("Validation failed: Row counts do not match.");
            }
        }
    }
}