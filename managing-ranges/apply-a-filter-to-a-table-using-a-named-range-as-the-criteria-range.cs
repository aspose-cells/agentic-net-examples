// Title: C# – Apply Advanced Filter to an Aspose.Cells Table Using a Named Range
// Description: Creates a workbook, builds a two‑column ListObject, defines a named range (CriteriaRange) with a matching header, and runs Worksheet.AdvancedFilter to filter the table in place. The filtered file is saved as FilteredTableWithNamedCriteria.xlsx.
// Keywords: Aspose.Cells | C# | AdvancedFilter | named range | criteria range | ListObject filter | Excel table filtering | in‑place filter | Aspose.Cells example | Excel automation .NET
// Common Searches: Aspose.Cells advanced filter with named range C# | How to filter a ListObject using a criteria range in Aspose.Cells | Apply in‑place AdvancedFilter to an Excel table with Aspose.Cells | Create and use named range for Excel filter Aspose.Cells .NET | C# code sample for Aspose.Cells AdvancedFilter
// Developer Intent: Filter rows of an Excel table programmatically by applying Aspose.Cells' AdvancedFilter with a named criteria range.
// Use Cases: Show only expense rows where Category equals "Food" without copying data to another range. | Reuse a single named criteria range to filter multiple tables across a workbook. | Perform quick, in‑place data segmentation while keeping original headers and layout intact.
// AI Prompts: Generate C# code that creates a named range for filter criteria and applies Aspose.Cells AdvancedFilter to a ListObject. | Explain how to modify the named criteria range to filter different values (e.g., other categories) using Aspose.Cells. | Suggest robust error‑handling for cases where the named criteria range header does not match any table column.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace AsposeCellsFilterWithNamedRange
{
    // Creates a workbook, builds a two‑column ListObject, defines a named range (CriteriaRange) with a matching header, and runs Worksheet.AdvancedFilter to filter the table in place. The filtered file is saved as FilteredTableWithNamedCriteria.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Populate sample data (the table) ----------
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["A2"].PutValue("Food");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["A3"].PutValue("Transport");
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["A4"].PutValue("Food");
                worksheet.Cells["B4"].PutValue(150);

                // Create a ListObject (table) from the data range A1:B4
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int listObjectIndex = worksheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = worksheet.ListObjects[listObjectIndex];
                // No need to set ShowHeaders; it is already true when hasHeaders = true

                // ---------- Define criteria range ----------
                // Criteria header (must match a column header) and the filter value
                worksheet.Cells["D1"].PutValue("Category");
                worksheet.Cells["D2"].PutValue("Food");

                // Create a named range that refers to the criteria cells (D1:D2)
                AsposeRange criteriaRange = worksheet.Cells.CreateRange("D1:D2");
                criteriaRange.Name = "CriteriaRange";

                // ---------- Apply Advanced Filter ----------
                // isFilter = false  -> filter in place (do not copy to another range)
                // listRange  = address of the table data (including headers)
                // criteriaRange = name of the criteria range defined above
                // copyTo = null (not used when filtering in place)
                // uniqueRecordOnly = false (show all matching rows)
                worksheet.AdvancedFilter(
                    false,                 // filter in place
                    "A1:B4",               // list range (table)
                    "CriteriaRange",       // criteria range (named range)
                    null,                  // copy to range (not needed)
                    false);                // do not restrict to unique records

                // Save the workbook to verify the filter was applied
                workbook.Save("FilteredTableWithNamedCriteria.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
