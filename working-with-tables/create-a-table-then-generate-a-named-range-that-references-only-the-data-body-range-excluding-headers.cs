// Title: Aspose.Cells for .NET: Create an Excel table and name only its data body range (C#)
// Description: Demonstrates how to build a new workbook, add a ListObject (table) with headers, extract the table's DataRange that excludes the header row, assign a custom name to that range, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# named range | Excel table DataRange Aspose | ListObject DataRange without header | create named range from table body | .NET Excel table named range | Aspose.Cells exclude header range
// Common Searches: Aspose.Cells create named range for table data only | C# get table DataRange without header Aspose | how to name only the data rows of an Excel table using Aspose.Cells | Aspose.Cells ListObject DataRange example | named range for Excel table body C#
// Developer Intent: Generate a named range that points exclusively to the data rows of a ListObject, omitting the header, with Aspose.Cells for .NET.
// Use Cases: Reference table rows in formulas, charts, or pivot tables without the header. | Export or import a clean data block to external systems that require header‑free ranges. | Apply data validation, conditional formatting, or VBA macros to the table body via a named range.
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject, retrieve its DataRange, and create a named range called 'EmployeeData' that excludes the header. | Explain the difference between ListObject.Range and ListObject.DataRange in Aspose.Cells and when to use each for named ranges. | Provide a step‑by‑step guide to verify the created named range in Excel after saving the workbook.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTableNamedRangeDemo
{
    // Demonstrates how to build a new workbook, add a ListObject (table) with headers, extract the table's DataRange that excludes the header row, assign a custom name to that range, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Register code page provider (required for some locales)
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with a header row
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("John");
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("Mary");
                worksheet.Cells["A4"].PutValue(3);
                worksheet.Cells["B4"].PutValue("Steve");

                // Add a ListObject (table) that includes headers
                // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "EmployeeTable";

                // Retrieve the data range of the table (excludes header row)
                AsposeRange dataRange = table.DataRange;

                // Create a named range that points to the data body only
                dataRange.Name = "EmployeeData";

                // Save the workbook
                workbook.Save("EmployeeTableWithNamedDataRange.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
