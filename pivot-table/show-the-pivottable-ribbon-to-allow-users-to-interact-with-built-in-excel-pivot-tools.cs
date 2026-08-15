// Title: Show Excel PivotTable Ribbon (Field List, Dialog, Wizard) with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills it with sample sales data, adds a PivotTable at cell E3, assigns row and data fields, and activates the built‑in PivotTable Field List, Field Dialog, and Wizard ribbon features. The pivot cache is refreshed and calculated before saving the file as ShowPivotTableRibbonDemo.xlsx.
// Keywords: Aspose.Cells PivotTable ribbon | C# enable PivotTable field list | Aspose.Cells show PivotTable wizard | Excel PivotTable UI features .NET | EnableFieldList Aspose.Cells | EnableFieldDialog C# | EnableWizard Aspose.Cells | refresh calculate pivot Aspose.Cells | display PivotTable ribbon tools
// Common Searches: how to show pivot table field list ribbon using Aspose.Cells C# | enable pivot table wizard in generated Excel with Aspose.Cells | Aspose.Cells refresh and calculate pivot table after UI enable | display built‑in Excel pivot tools in .NET workbook | C# code to activate PivotTable field dialog with Aspose.Cells
// Developer Intent: Display the built‑in PivotTable ribbon tools (Field List, Field Dialog, Wizard) in an Excel file generated with Aspose.Cells for .NET.
// Use Cases: Generate an Excel workbook that end users can modify via the PivotTable ribbon UI. | Programmatically turn on the Field List, Field Dialog, and Wizard so users can rearrange rows, columns, and values after opening the file. | Ensure pivot data is up‑to‑date by refreshing and calculating the cache before saving.
// AI Prompts: Write C# code using Aspose.Cells to add a PivotTable and enable the Field List, Field Dialog, and Wizard ribbon features. | Explain the effect of EnableFieldList, EnableFieldDialog, and EnableWizard on the Excel UI and any required steps before saving the workbook. | Provide a step‑by‑step guide to refresh and calculate a PivotTable after enabling UI features with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills it with sample sales data, adds a PivotTable at cell E3, assigns row and data fields, and activates the built‑in PivotTable Field List, Field Dialog, and Wizard ribbon features. The pivot cache is refreshed and calculated before saving the file as ShowPivotTableRibbonDemo.xlsx.
    public class ShowPivotTableRibbonDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Product";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Electronics";
                sheet.Cells["B2"].Value = "Laptop";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Electronics";
                sheet.Cells["B3"].Value = "Phone";
                sheet.Cells["C3"].Value = 800;

                sheet.Cells["A4"].Value = "Furniture";
                sheet.Cells["B4"].Value = "Chair";
                sheet.Cells["C4"].Value = 150;

                sheet.Cells["A5"].Value = "Furniture";
                sheet.Cells["B5"].Value = "Table";
                sheet.Cells["C5"].Value = 300;

                // Add a pivot table to a new location (E3)
                int pivotIndex = sheet.PivotTables.Add("=A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Enable built‑in PivotTable UI ribbon features
                pivotTable.EnableFieldList = true;    // Shows the PivotTable Field List ribbon
                pivotTable.EnableFieldDialog = true; // Allows double‑click to open the field dialog
                pivotTable.EnableWizard = true;      // Makes the PivotTable Wizard accessible

                // Refresh and calculate the pivot data using the correct API
                pivotTable.RefreshData();   // Refreshes the pivot cache
                pivotTable.CalculateData(); // Calculates the pivot table values

                // Save the workbook
                string outputPath = "ShowPivotTableRibbonDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShowPivotTableRibbonDemo.Run();
        }
    }
}
