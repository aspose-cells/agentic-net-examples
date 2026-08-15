// Title: Hide PivotTable Field List Ribbon at Runtime with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add sample data, build a PivotTable on a separate sheet, and disable the PivotTable field‑list ribbon using Workbook.Settings.HidePivotFieldList so the Excel file opens with a clean interface.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | HidePivotFieldList | hide pivot field list | disable pivot ribbon | runtime hide pivot UI | Excel pivot table UI | Aspose.Cells API | Workbook.Settings.HidePivotFieldList
// Common Searches: Aspose.Cells hide PivotTable field list ribbon | disable PivotTable UI in generated Excel using C# | runtime hide PivotTable ribbon Aspose.Cells | how to suppress PivotTable field list with Aspose.Cells | remove PivotTable ribbon interface programmatically
// Developer Intent: Programmatically suppress the PivotTable field‑list ribbon so the generated workbook opens without the PivotTable UI.
// Use Cases: Create a reporting workbook where the PivotTable layout must stay unchanged for end users. | Generate a dashboard Excel file that displays a clean view without the PivotTable controls. | Distribute a template with a pre‑configured PivotTable while preventing users from accessing the field list.
// AI Prompts: Show C# code to hide the PivotTable field list ribbon in an Aspose.Cells workbook. | Provide an example that creates a PivotTable and disables its UI at runtime using Aspose.Cells. | Explain the effect of Workbook.Settings.HidePivotFieldList and how to re‑enable the field list later.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add sample data, build a PivotTable on a separate sheet, and disable the PivotTable field‑list ribbon using Workbook.Settings.HidePivotFieldList so the Excel file opens with a clean interface.
    public class HidePivotTableRibbonDemo
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(150);

            // Create a separate worksheet for the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add(SheetType.Worksheet);
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotTable";

            // Add the pivot table
            int pivotTableIndex = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Value as data field

            // Hide the PivotTable field list (ribbon interface) at runtime
            workbook.Settings.HidePivotFieldList = true;

            // Save the workbook
            workbook.Save("HidePivotTableRibbonDemo.xlsx");
        }
    }
}
