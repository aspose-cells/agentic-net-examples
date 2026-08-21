// Title: C# – Apply Theme Accent2 Fill to Pivot Table Summary Cells with Aspose.Cells
// Description: Creates a workbook, adds sample sales data, builds a pivot table, defines a style that uses the workbook's Theme Accent2 color as a solid background, applies the style to the pivot table's DataBodyRange (summary cells), and saves the file.
// Keywords: Aspose.Cells | C# | PivotTable formatting | Theme Accent2 | BackgroundThemeColor | DataBodyRange | Excel report styling | solid fill | theme color in code | pivot table summary cells
// Common Searches: Aspose.Cells set theme Accent2 background for pivot table summary | C# apply solid fill to pivot table data area using theme color | How to format pivot table summary cells with Accent2 in Aspose.Cells | Apply theme color to pivot table values .NET | PivotTable DataBodyRange style Aspose.Cells C#
// Developer Intent: Use the workbook's Accent2 theme color to fill the pivot table's summary cells.
// Use Cases: Highlight pivot table totals with the corporate Accent2 theme for consistent branding. | Generate automated Excel reports where summarized values are visually distinguished by a theme‑based fill. | Create sales dashboards that apply a solid Accent2 background to pivot data cells for quick data interpretation.
// AI Prompts: Show how to change the fill to Theme Accent3 for the pivot table summary cells. | Explain how to add a 30% tint to the Accent2 background when formatting pivot data. | Provide code that applies the Accent2 fill to both the data body and grand total rows of a pivot table.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotThemeAccent2
{
    // Creates a workbook, adds sample sales data, builds a pivot table, defines a style that uses the workbook's Theme Accent2 color as a solid background, applies the style to the pivot table's DataBodyRange (summary cells), and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Food";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 850;

            sheet.Cells["A4"].Value = "Drink";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 560;

            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 730;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the pivot data
            pivot.CalculateData();

            // Create a style that uses the theme's Accent2 color for the background fill
            Style accentStyle = workbook.CreateStyle();
            accentStyle.Pattern = BackgroundType.Solid;                     // Solid fill
            accentStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0); // Accent2, no tint

            // Apply the style to the data (summary) area of the pivot table
            // DataBodyRange represents the cells that contain the summarized values
            CellArea dataArea = pivot.DataBodyRange;
            pivot.Format(dataArea, accentStyle);

            // Save the workbook
            workbook.Save("PivotTable_Accent2_Summary.xlsx");
        }
    }
}
