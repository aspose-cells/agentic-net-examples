// Title: Apply a Theme Color to PivotTable Headers with FormatAll in Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, sets the Accent1 theme color to orange, builds a Style that uses the theme color for the header font and a light‑yellow fill, adds a PivotTable, and applies the style to every PivotTable cell—including all header rows—using the FormatAll method before saving the file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | header style | theme color | Accent1 | FormatAll | Excel report | style object | sample code | branding
// Common Searches: Aspose.Cells format all pivot table headers C# | set theme color for pivot table headers Aspose.Cells | apply style to entire pivot table using FormatAll | change pivot table header font color with theme in .NET | sample code for PivotTable header styling Aspose.Cells
// Developer Intent: The developer wants to apply a predefined style that uses a custom theme color to all PivotTable headers in an Excel workbook.
// Use Cases: Generate Excel reports where PivotTable headers automatically follow the workbook’s branding colors. | Create a reusable Style object to format headers of multiple PivotTables with a single method call. | Automate consistent header appearance across dashboards that are built programmatically with Aspose.Cells.
// AI Prompts: Write C# code using Aspose.Cells to set Accent1 to orange and apply that theme color to all PivotTable headers with FormatAll. | Show how to create a Style with a theme‑based font and background, then apply it to a PivotTable in Aspose.Cells for .NET. | Explain how to reuse a single Style object to format headers of several PivotTables in the same workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, sets the Accent1 theme color to orange, builds a Style that uses the theme color for the header font and a light‑yellow fill, adds a PivotTable, and applies the style to every PivotTable cell—including all header rows—using the FormatAll method before saving the file.
class ApplyThemeColorToPivotHeaders
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["A2"].Value = "Food";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["A3"].Value = "Drink";
        sheet.Cells["B3"].Value = 150;
        sheet.Cells["A4"].Value = "Snack";
        sheet.Cells["B4"].Value = 80;

        // Define a custom theme color (Accent1) that will be used in the style
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.Orange);

        // Create a style that uses the theme color for the font (header appearance)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Size = 12;
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.ForegroundColor = Color.LightYellow;

        // Add a pivot table based on the sample data
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Apply the predefined style to all cells in the pivot table (including headers)
        pivot.FormatAll(headerStyle);

        // Save the workbook
        workbook.Save("PivotHeaderThemeDemo.xlsx", SaveFormat.Xlsx);
    }
}
