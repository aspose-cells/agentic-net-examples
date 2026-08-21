// Title: Aspose.Cells for .NET – Set Dark1 Theme Fill on Pivot Table Row Headers (C#)
// Description: Shows how to create a workbook, add sample data, build a pivot table, enable row‑header styling, apply the PivotTableStyleDark1 (which uses the workbook’s Dark1 theme color for fills), and save the file as an .xlsx document.
// Keywords: Aspose.Cells | C# pivot table | PivotTableStyleDark1 | Dark1 theme color | row header fill | Excel theme color | Aspose.Cells pivot styling | set pivot table row header color | Excel theme fill | Aspose.Cells .NET
// Common Searches: Aspose.Cells set Dark1 theme for pivot table row headers | C# apply PivotTableStyleDark1 to row headers | change pivot table row header fill color using Aspose.Cells | how to use workbook theme colors in Aspose.Cells pivot tables | apply theme color to pivot table headers .NET
// Developer Intent: Apply the workbook’s Dark1 theme color to the fill of pivot table row headers.
// Use Cases: Generate a formatted report where pivot table row headers match the workbook theme. | Standardize visual style of Excel dashboards created programmatically. | Ensure consistent theme‑aware coloring across multiple pivot tables in automated reporting. | Create pivot tables with theme‑based styling for corporate branding.
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table and sets the row header fill to the Dark1 theme color. | Explain the effect of PivotTableStyleDark1 on row header colors in an Aspose.Cells pivot table. | Show how to modify an existing Aspose.Cells pivot table to use the Dark1 theme for row header fills. | Provide step‑by‑step instructions to enable ShowPivotStyleRowHeader and apply PivotTableStyleDark1 in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add sample data, build a pivot table, enable row‑header styling, apply the PivotTableStyleDark1 (which uses the workbook’s Dark1 theme color for fills), and save the file as an .xlsx document.
    public class PivotTableRowHeaderDark1Theme
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "A";
                sheet.Cells["B2"].Value = 100;
                sheet.Cells["A3"].Value = "B";
                sheet.Cells["B3"].Value = 200;
                sheet.Cells["A4"].Value = "A";
                sheet.Cells["B4"].Value = 150;
                sheet.Cells["A5"].Value = "B";
                sheet.Cells["B5"].Value = 250;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: add a row field and a data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Ensure the row header style is applied
                pivotTable.ShowPivotStyleRowHeader = true;

                // Apply the Dark1 pivot table style – this uses the theme's Dark1 color for fills
                pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleDark1;

                // Save the workbook
                workbook.Save("PivotTableRowHeaderDark1.xlsx");
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
            PivotTableRowHeaderDark1Theme.Run();
        }
    }
}
