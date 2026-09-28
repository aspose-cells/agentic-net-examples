// Title: Enable multiple item selection in a PivotTable filter field with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets the IsMultipleItemSelectionAllowed flag on a PivotField to allow multi‑select filtering in an Aspose.Cells PivotTable. | Show how to create a workbook, add a pivot table, place a field in the page area, and enable multiple selection for that filter using the Aspose.Cells API. | Provide an Aspose.Cells example that configures a PivotTable’s page field to support selecting multiple items in the filter dialog. | Write a step‑by‑step guide to activate multi‑select on a PivotTable filter in a .NET workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# enable multi select on pivot table page field | How to allow multiple items selection in pivot table filter using Aspose.Cells .NET | Set IsMultipleItemSelectionAllowed property for PivotField Aspose.Cells example | C# code sample for multi‑select filter in Aspose.Cells PivotTable | Aspose.Cells enable multiple selection in pivot table filter dialog
// Tags: Aspose.Cells enable multiple selection pivot filter | C# set PivotField IsMultipleItemSelectionAllowed | Aspose.Cells create pivot table with page field | Aspose.Cells multi‑select filter dialog | Aspose.Cells PivotTable page field settings

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotMultipleSelectionDemo
{
    // Creates a workbook, adds sample data, builds a PivotTable, puts the "Fruit" field in the page (filter) area, sets the PivotField.IsMultipleItemSelectionAllowed property to true to enable multi‑select filtering, and saves the file as PivotTable_MultipleSelection.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 15;
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 20;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the field to the Page area (filter area)
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Fruit");

            // Enable multiple item selection for the page field
            PivotField pageField = pivotTable.PageFields[0];
            pageField.IsMultipleItemSelectionAllowed = true;

            // Optional: verify the property was set
            Console.WriteLine("IsMultipleItemSelectionAllowed: " + pageField.IsMultipleItemSelectionAllowed);

            // Save the workbook
            workbook.Save("PivotTable_MultipleSelection.xlsx");
        }
    }
}
