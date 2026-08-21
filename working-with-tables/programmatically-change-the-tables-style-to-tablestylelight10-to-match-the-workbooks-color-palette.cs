// Title: C# – Apply TableStyleLight10 to an Aspose.Cells ListObject (Table)
// Description: Shows how to create a workbook, insert sample data, define a ListObject, set its TableStyleType to the built‑in TableStyleLight10, and save the file so the table follows the workbook’s color palette.
// Keywords: Aspose.Cells | C# | .NET | ListObject | TableStyleLight10 | table style | built‑in table style | Excel table formatting | workbook palette
// Common Searches: Aspose.Cells set table style C# | TableStyleLight10 Aspose.Cells example | change ListObject style .NET | apply built‑in table style Aspose.Cells | C# code for Excel table formatting
// Developer Intent: Apply the built‑in TableStyleLight10 to a ListObject so the table inherits the workbook’s default color scheme.
// Use Cases: Generate Excel reports where every table uses a consistent light theme. | Automate workbook creation with tables that automatically adopt a predefined built‑in style. | Export data files that must match the workbook’s default color palette for branding compliance.
// AI Prompts: Provide C# code to change a ListObject’s TableStyleType to any built‑in style in Aspose.Cells. | List all TableStyleType enum values and show how to select one based on a condition. | Explain how to ensure a chosen table style aligns with the workbook’s theme colors in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to create a workbook, insert sample data, define a ListObject, set its TableStyleType to the built‑in TableStyleLight10, and save the file so the table follows the workbook’s color palette.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the table
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(25);

        // Create a ListObject (table) covering the data range
        int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Apply the built‑in style TableStyleLight10 to the table
        table.TableStyleType = TableStyleType.TableStyleLight10;

        // Save the workbook
        workbook.Save("TableStyleLight10.xlsx");
    }
}
