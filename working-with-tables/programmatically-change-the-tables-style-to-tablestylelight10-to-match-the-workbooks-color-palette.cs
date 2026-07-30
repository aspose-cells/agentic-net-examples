// Title: Aspose.Cells C# Example: Apply TableStyleLight10 to a ListObject Table
// Description: Demonstrates how to create a workbook, add data, convert the range to a ListObject, set its TableStyleType to the built‑in TableStyleLight10 (matching the workbook palette), and save the file as XLSX.
// Keywords: Aspose.Cells | C# | ListObject | TableStyleLight10 | Excel table style | apply built‑in style | programmatic table formatting | Aspose.Cells API | TableStyleType | Excel automation
// Common Searches: Aspose.Cells set table style C# | TableStyleLight10 ListObject Aspose | apply built‑in Excel table style with Aspose.Cells | change Aspose.Cells table style programmatically | C# code for Aspose.Cells table styling
// Developer Intent: Set the built‑in TableStyleLight10 on a ListObject table in an Aspose.Cells workbook.
// Use Cases: Generate a new Excel file with a styled table that follows the default workbook theme. | Standardize the appearance of existing tables across worksheets by applying TableStyleLight10. | Batch process multiple workbooks to enforce a consistent table style for corporate branding.
// AI Prompts: Provide C# code using Aspose.Cells to change a ListObject's TableStyleType to TableStyleLight10. | Show how to iterate over all tables in a worksheet and assign TableStyleLight10 in Aspose.Cells. | Explain the relationship between TableStyleLight10 and the workbook's color palette in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a workbook, add data, convert the range to a ListObject, set its TableStyleType to the built‑in TableStyleLight10 (matching the workbook palette), and save the file as XLSX.
class ChangeTableStyle
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the table
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(25);
        worksheet.Cells["A4"].PutValue("Doe");
        worksheet.Cells["B4"].PutValue(40);

        // Add a ListObject (table) that includes the data range
        // Parameters: first row, first column, last row, last column, hasHeaders
        int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Change the table's style to the built‑in style "TableStyleLight10"
        table.TableStyleType = TableStyleType.TableStyleLight10;

        // Optionally, ensure the style name reflects the change (not required but illustrative)
        // table.TableStyleName = "TableStyleLight10";

        // Save the workbook to a file
        workbook.Save("TableWithStyleLight10.xlsx");
    }
}
