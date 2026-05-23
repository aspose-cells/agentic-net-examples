using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFilterWithNamedRange
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Populate sample data (the list range) -----
                // Headers
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");

                // Data rows
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Transport");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["B4"].PutValue(150);

                // ----- Populate criteria range (the filter criteria) -----
                // Header for criteria (must match the list header)
                sheet.Cells["D1"].PutValue("Category");
                // Desired criteria value
                sheet.Cells["D2"].PutValue("Food");

                // Create a named range that refers to the criteria cells
                AsposeRange criteriaNamedRange = sheet.Cells.CreateRange("D1:D2");
                criteriaNamedRange.Name = "MyCriteria";

                // Apply an advanced filter:
                //   isFilter = false  -> filter in place (hide non‑matching rows)
                //   listRange = "A1:B4"  -> the data to be filtered
                //   criteriaRange = "MyCriteria" -> the named range defined above
                //   copyTo = null -> no copy, just filter in place
                //   uniqueRecordOnly = false -> keep all matching rows
                sheet.AdvancedFilter(false, "A1:B4", "MyCriteria", null, false);

                // Save the workbook to verify the filter was applied
                string outputPath = "FilteredTableWithNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Filter applied using named range successfully. File saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}