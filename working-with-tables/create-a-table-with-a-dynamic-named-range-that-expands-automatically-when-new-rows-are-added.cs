// Title: C# – Create a Dynamic Named Range that Grows with an Excel Table using Aspose.Cells
// Description: This C# example demonstrates how to add a ListObject (Excel table) to a workbook, define a named range that points to =EmployeeTable[#All], insert new rows, resize the table, and verify that the named range automatically expands before saving the file.
// Keywords: Aspose.Cells | C# dynamic named range | Excel table auto expand | ListObject resize | named range #All | Aspose.Cells .NET example | auto‑growing range | Excel named range formula | Aspose.Cells API
// Common Searches: Aspose.Cells create dynamic named range | C# expand named range with table rows | How to resize ListObject in Aspose.Cells | Get address of named range after adding rows | Reference Excel table in named range Aspose.Cells
// Developer Intent: The developer needs a named range that automatically includes new rows added to an Excel table.
// Use Cases: Use the dynamic range in formulas, charts, or pivot tables so they always cover the full employee list. | Apply data validation, conditional formatting, or data bars to an expanding range without manual updates. | Export the workbook to PDF, CSV, or other formats while preserving the auto‑growing range for downstream processing. | Integrate the range with reporting tools that rely on a stable named range identifier.
// AI Prompts: Write C# code using Aspose.Cells to create a ListObject and a named range that references EmployeeTable[#All] and expands automatically when rows are added. | Show how to add multiple rows to the table, resize it, and output the named range address before and after the insertion. | Explain how to reference the dynamic named range in worksheet formulas, charts, or pivot tables with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// This C# example demonstrates how to add a ListObject (Excel table) to a workbook, define a named range that points to =EmployeeTable[#All], insert new rows, resize the table, and verify that the named range automatically expands before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with headers
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");

            // Add a table (ListObject) covering the initial data range
            int tableIdx = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[tableIdx];
            table.DisplayName = "EmployeeTable";

            // Create a dynamic named range that refers to the whole table.
            // The reference "=EmployeeTable[#All]" expands automatically as the table grows.
            int nameIdx = workbook.Worksheets.Names.Add("DynamicEmployees");
            Name dynamicName = workbook.Worksheets.Names[nameIdx];
            dynamicName.RefersTo = "=EmployeeTable[#All]";

            // Show the address of the named range before adding new rows
            AsposeRange rangeBefore = dynamicName.GetRange();
            Console.WriteLine("Named range before adding rows: " + rangeBefore.Address);

            // Add a new row to the worksheet (below the current table)
            AsposeRange dataRange = table.DataRange;
            int newRowIndex = dataRange.FirstRow + dataRange.RowCount; // index of the new row (0‑based)

            // Fill data in the newly added row
            worksheet.Cells[newRowIndex, 0].PutValue(3);
            worksheet.Cells[newRowIndex, 1].PutValue("Bob");

            // Resize the table to include the new row (hasHeaders = true)
            table.Resize(dataRange.FirstRow, dataRange.FirstColumn, dataRange.RowCount + 1, dataRange.ColumnCount, true);

            // Retrieve the named range again to demonstrate that it has expanded
            AsposeRange rangeAfter = dynamicName.GetRange();
            Console.WriteLine("Named range after adding rows: " + rangeAfter.Address);

            // Save the workbook
            workbook.Save("DynamicTableNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
