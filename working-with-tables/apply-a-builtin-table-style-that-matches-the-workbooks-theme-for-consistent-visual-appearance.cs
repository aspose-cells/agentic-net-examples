// Title: Apply a built‑in, theme‑aware table style to a ListObject in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, insert sample data, define a ListObject for the data range, and set its TableStyleType to TableStyleMedium2 to apply the built‑in theme‑aware style with Aspose.Cells for .NET. | Load an existing worksheet, locate a ListObject, change its TableStyleType to TableStyleMedium2, ensure the header row is visible, and save the workbook as an .xlsx file.
// Common Searches: how to set a built‑in table style that follows the workbook theme in Aspose.Cells C# | Aspose.Cells C# change ListObject style to TableStyleMedium2 | apply theme‑aware table formatting to Excel using Aspose.Cells .NET | C# add a table with header row and apply built‑in style in Aspose.Cells | save Excel workbook with styled ListObject using Aspose.Cells
// Tags: built‑in table style application Aspose.Cells | theme‑aware ListObject styling C# | TableStyleMedium2 usage Excel | create ListObject with header Aspose | export styled workbook to .xlsx .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a workbook, populates product data, adds a ListObject covering the range A1:C4, assigns the built‑in theme‑aware TableStyleMedium2, ensures the header row is shown, and saves the result as AppliedTableStyle.xlsx.
class ApplyTableStyle
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["C4"].PutValue(1.2);

            // Define the range for the table (zero‑based indices)
            int firstRow = 0;      // Row 1
            int firstColumn = 0;   // Column A
            int totalRows = 4;     // Including header row
            int totalColumns = 3;

            // Add a ListObject (table) with header row flag set to true
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "ProductTable";

            // Apply a built‑in, theme‑aware table style
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Optional: show header row (default is true)
            table.ShowHeaderRow = true;
            // Total row is hidden by default; no need to set ShowTotalRow

            // Save the workbook
            string outputPath = "AppliedTableStyle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
