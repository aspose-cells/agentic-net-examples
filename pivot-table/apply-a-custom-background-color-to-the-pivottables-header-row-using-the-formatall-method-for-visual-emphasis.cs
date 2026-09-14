// Title: Apply a custom background color to a PivotTable header row with Aspose.Cells FormatAll in C#
// AI Prompts: Generate C# code that creates a workbook, builds a PivotTable from a data range, defines a Style with a solid LightGoldenrodYellow background and bold font, and calls PivotTable.FormatAll to apply the style to the table, emphasizing the header row. | Show how to adjust the Aspose.Cells example to use a different background color and font weight for the PivotTable header while leaving the rest of the table unchanged. | Explain step‑by‑step how the Style object and FormatAll method work together to format PivotTable headers in an XLSX file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# change pivot table header background color | How to use FormatAll to style pivot table headers in .NET | Set solid background and bold font for pivot table header row with Aspose.Cells | Apply custom style to pivot table header only using Aspose.Cells API | C# example formatting pivot table header row in Excel workbook
// Tags: pivot table header background styling Aspose.Cells | FormatAll method style application C# | custom solid background color Excel pivot header | Aspose.Cells style object for pivot tables | C# Excel workbook pivot table formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotHeaderFormatting
{
    // The example creates a new workbook, adds sample data, builds a PivotTable, defines a Style with a LightGoldenrodYellow solid background and bold font, applies the style to the entire PivotTable using the FormatAll method (which highlights the header row), and saves the result as PivotTableHeaderFormatted.xlsx.
    class Program
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
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Drink";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Food";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Sum of Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Create a style with a custom background color for emphasis
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightGoldenrodYellow; // custom background color
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.IsBold = true; // make header text bold for extra emphasis

            // Apply the style to the entire pivot table using FormatAll
            // (this demonstrates the required method; it will affect all cells,
            // including the header row, providing the visual emphasis requested)
            pivotTable.FormatAll(headerStyle);

            // Save the workbook
            workbook.Save("PivotTableHeaderFormatted.xlsx", SaveFormat.Xlsx);
        }
    }
}
