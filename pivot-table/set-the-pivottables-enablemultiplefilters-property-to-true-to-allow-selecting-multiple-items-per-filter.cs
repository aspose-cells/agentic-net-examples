// Title: Aspose.Cells C# – Enable Multi‑Select Filters in a PivotTable (AllowMultipleFiltersPerField = true)
// Description: This C# example creates a workbook, fills a small data set, adds a PivotTable, assigns the Category field to rows and the Amount field to values, then sets AllowMultipleFiltersPerField to true so users can select multiple items in each filter. The workbook is saved as PivotTable_MultipleFilters.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# PivotTable | AllowMultipleFiltersPerField | EnableMultipleFilters | multi‑select pivot filter | multiple filters pivot table | Excel pivot table code example | Aspose.Cells pivot filter settings
// Common Searches: Aspose.Cells enable multiple filters pivot table C# | AllowMultipleFiltersPerField example Aspose.Cells | How to allow multi‑select filters in Aspose.Cells PivotTable | Set EnableMultipleFilters property Aspose.Cells .NET | C# code for PivotTable with multiple filters Aspose
// Developer Intent: Show how to activate multi‑select filtering for each field in an Aspose.Cells PivotTable using the AllowMultipleFiltersPerField property.
// Use Cases: Sales analysis where users need to filter by several product categories simultaneously. | Interactive dashboard that lets analysts pick multiple regions or time periods in a pivot filter. | Automated report generation that requires a pivot with multi‑select filters for flexible data slicing.
// AI Prompts: Provide a C# snippet that creates an Aspose.Cells PivotTable and enables multi‑select filters with AllowMultipleFiltersPerField. | Explain the impact of AllowMultipleFiltersPerField on filter behavior in an Aspose.Cells PivotTable. | Generate code to add a PivotTable from a range and turn on multiple filters per field using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotMultipleFiltersDemo
{
    // This C# example creates a workbook, fills a small data set, adds a PivotTable, assigns the Category field to rows and the Amount field to values, then sets AllowMultipleFiltersPerField to true so users can select multiple items in each filter. The workbook is saved as PivotTable_MultipleFilters.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Drink";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Food";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Drink";
            cells["B5"].Value = 60;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable multiple filters per field
            pivotTable.AllowMultipleFiltersPerField = true;

            // Save the workbook
            workbook.Save("PivotTable_MultipleFilters.xlsx");

            // Optional: output confirmation
            Console.WriteLine("Pivot table created with AllowMultipleFiltersPerField = true.");
        }
    }
}
