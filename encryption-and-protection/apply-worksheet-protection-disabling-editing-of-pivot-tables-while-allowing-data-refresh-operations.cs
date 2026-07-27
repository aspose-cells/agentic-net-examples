// Title: C# – Protect an Excel worksheet, disable pivot table editing while allowing refresh with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a pivot table, set worksheet protection to block pivot table modifications (AllowUsingPivotTable = false), keep refresh operations enabled, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection C# | disable pivot table editing Aspose | allow pivot refresh protected sheet | ProtectionType.All Aspose.Cells | C# Excel pivot table protection example
// Common Searches: Aspose.Cells protect sheet but allow pivot refresh | C# disable pivot table editing on protected worksheet | How to set AllowUsingPivotTable false in Aspose.Cells | Refresh pivot tables after worksheet protection Aspose | Excel worksheet protection with pivot tables C#
// Developer Intent: Protect a worksheet, prevent users from altering pivot tables, yet still permit pivot table refresh operations.
// Use Cases: Distribute a reporting template where users can update data via refresh but cannot change the pivot layout. | Secure shared workbooks in a corporate environment while keeping automated data refresh functional. | Create dashboards that lock pivot configurations but allow real‑time data updates.
// AI Prompts: Show C# code to protect an Excel sheet with Aspose.Cells, disable pivot table editing, and still allow pivot refresh. | Explain the impact of setting AllowUsingPivotTable = false together with ProtectionType.All in Aspose.Cells. | Generate a step‑by‑step guide for protecting a worksheet while keeping pivot table refresh enabled using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a pivot table, set worksheet protection to block pivot table modifications (AllowUsingPivotTable = false), keep refresh operations enabled, and save the file using Aspose.Cells for .NET.
    public class ProtectWorksheetDisablePivotEditingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruit");
            worksheet.Cells["A3"].PutValue("Vegetable");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(80);

            // Add a pivot table based on the sample data
            int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D5", "SamplePivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity as data field

            // Access the worksheet protection settings
            Protection protection = worksheet.Protection;

            // Disable manipulation of pivot tables while the sheet is protected
            protection.AllowUsingPivotTable = false;

            // Protect the worksheet (all protection types, no password)
            worksheet.Protect(ProtectionType.All);

            // Refresh pivot tables – this operation is still allowed despite the protection setting
            worksheet.RefreshPivotTables();

            // Save the workbook
            workbook.Save("ProtectWorksheetDisablePivotEditingDemo.xlsx");
        }
    }
}
