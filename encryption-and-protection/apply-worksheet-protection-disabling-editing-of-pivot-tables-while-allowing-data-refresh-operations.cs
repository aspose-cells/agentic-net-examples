// Title: C# – Protect Worksheet to Block Pivot Table Editing While Allowing Refresh with Aspose.Cells
// Description: Shows how to build a workbook, insert a pivot table, and apply worksheet protection that disables pivot table editing (AllowUsingPivotTable = false) yet still permits RefreshPivotTables. The sheet is secured with ProtectionType.All (no password) and saved as an Excel file.
// Keywords: Aspose.Cells worksheet protection C# | disable pivot table editing | AllowUsingPivotTable false | refresh pivot on protected sheet | Aspose.Cells .NET example | Excel pivot lock without password | global Excel security | US developers Aspose.Cells
// Common Searches: Aspose.Cells protect worksheet but keep pivot refreshable | C# disable pivot table editing Aspose.Cells | AllowUsingPivotTable property example | RefreshPivotTables on protected sheet | How to lock pivot layout in Excel using Aspose.Cells
// Developer Intent: Secure a worksheet so users cannot modify pivot tables, yet they can still refresh the pivot data programmatically.
// Use Cases: Distribute a financial report where the pivot layout is locked but data can be refreshed by end‑users. | Create a dashboard template that prevents accidental pivot changes while allowing automated nightly refreshes. | Provide external partners with a protected Excel file that only permits pivot refresh, not structural edits.
// AI Prompts: Generate C# code using Aspose.Cells that protects a worksheet, disables pivot table editing, and still allows RefreshPivotTables. | Explain the effect of setting Protection.AllowUsingPivotTable = false together with ProtectionType.All in Aspose.Cells. | Show how to protect an Excel sheet without a password while keeping pivot tables refreshable in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to build a workbook, insert a pivot table, and apply worksheet protection that disables pivot table editing (AllowUsingPivotTable = false) yet still permits RefreshPivotTables. The sheet is secured with ProtectionType.All (no password) and saved as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruit");
        worksheet.Cells["A3"].PutValue("Vegetable");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);

        // Add a pivot table based on the sample data
        int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D5", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Configure worksheet protection:
        // - Disallow manipulation of pivot tables (editing, moving, etc.)
        // - Allow all other operations, including data refresh
        Protection protection = worksheet.Protection;
        protection.AllowUsingPivotTable = false; // Disable editing of pivot tables
        // Protect the worksheet (no password required for this example)
        worksheet.Protect(ProtectionType.All);

        // Refresh pivot tables – this operation is still permitted despite the protection setting
        worksheet.RefreshPivotTables();

        // Save the workbook
        workbook.Save("ProtectedWorksheetWithPivotRefresh.xlsx");
    }
}
