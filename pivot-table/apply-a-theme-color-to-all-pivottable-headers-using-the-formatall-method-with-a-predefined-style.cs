// Title: Apply a workbook theme accent color to all PivotTable headers using PivotTable.FormatAll in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a PivotTable from a data range, defines a custom workbook theme accent, builds a style using that theme, and applies the style to the PivotTable headers with the FormatAll method. | Demonstrate how to set a custom Accent1 theme color in an Aspose.Cells workbook and use it to format PivotTable header fonts and background via a ThemeColor‑based style. | Provide a complete example that formats PivotTable headers with a bold orange font and light‑yellow fill by creating a style that references the workbook theme and applying it through PivotTable.FormatAll.
// Common Searches: Aspose.Cells C# apply workbook theme color to pivot table header style | How to use PivotTable.FormatAll to style headers in .NET | Set custom Accent1 theme color and apply to PivotTable in Aspose.Cells | Formatting pivot table header font color with theme in C# Aspose.Cells
// Tags: pivot table header styling with FormatAll Aspose.Cells | apply workbook theme color to pivot table Aspose.Cells .NET | custom theme accent for pivot table headers C# | style entire pivot table using Aspose.Cells API | theme‑based style for pivot table headers XLSX

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotThemeDemo
{
    // The example creates a new workbook, adds sample data, builds a PivotTable, sets the workbook's Accent1 theme color to orange, creates a style that uses this theme color for a bold 12‑point font with a light‑yellow background, applies the style to the entire PivotTable (including headers) via the FormatAll method, and saves the file as an XLSX workbook.
    public class ApplyThemeToPivotHeaders
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
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Food";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["A3"].Value = "Drink";
                sheet.Cells["B3"].Value = 80;
                sheet.Cells["A4"].Value = "Food";
                sheet.Cells["B4"].Value = 150;
                sheet.Cells["A5"].Value = "Drink";
                sheet.Cells["B5"].Value = 70;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Configure the pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Define a custom theme color (Accent1) that will be used in the style
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.Orange);

                // Create a style that uses the theme color for the font
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
                headerStyle.Font.IsBold = true;
                headerStyle.Font.Size = 12;
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.ForegroundColor = Color.LightYellow;

                // Apply the style to the entire pivot table area (including headers)
                pivot.FormatAll(headerStyle);

                // Save the workbook
                workbook.Save("PivotTableWithThemeHeaders.xlsx", SaveFormat.Xlsx);
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
            ApplyThemeToPivotHeaders.Run();
        }
    }
}
