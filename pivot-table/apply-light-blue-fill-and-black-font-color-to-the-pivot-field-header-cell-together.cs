// Title: How to apply a light‑blue fill and black font to a pivot table row field header using Aspose.Cells for .NET (C#)
// AI Prompts: Define a style with a solid light‑blue background and black font, then assign it to the row field header of a pivot table using PivotFormats.FormatArea in Aspose.Cells C#. | Use Workbook.CreateStyle to set background and font colors, and format the pivot field header cell of a generated pivot table.
// Common Searches: how to change pivot table row header background color in Aspose.Cells C# | formatting pivot field header font color with Aspose.Cells .NET | example of applying custom style to pivot table header cell in C# | Aspose.Cells PivotFormats.FormatArea to style pivot table headers | set solid fill for pivot table header using Aspose.Cells API
// Tags: Aspose.Cells pivot table header styling | C# PivotFormats.FormatArea example | solid fill for pivot field header | set font color on pivot table header .NET | Workbook.CreateStyle background color Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHeaderFormatting
{
    // The program creates a workbook, adds sample data, builds a pivot table, defines a style with a light‑blue solid fill and black font, applies this style to the row field header cell via PivotFormats.FormatArea, and saves the file as PivotHeaderFormatted.xlsx.
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

            // Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields: Category as row field, Amount as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Calculate the pivot data
            pivotTable.CalculateData();

            // Create a style: light blue fill and black font color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.BackgroundColor = Color.LightBlue;          // Fill color
            headerStyle.Pattern = BackgroundType.Solid;            // Apply fill
            headerStyle.Font.Color = Color.Black;                  // Font color
            headerStyle.Font.IsBold = true;                        // Optional: make it bold

            // Apply the style to the row header area (pivot field header cell)
            // axisType = Row, fieldPosition = 0 (first row field), no subtotal,
            // selection includes both data and label, not grand totals.
            pivotTable.PivotFormats.FormatArea(
                PivotFieldType.Row,          // axis type
                0,                           // field position (first row field)
                PivotFieldSubtotalType.None,// subtotal type
                PivotTableSelectionType.DataAndLabel, // select header cell
                false,                       // isGrandRow
                false,                       // isGrandColumn
                headerStyle);                // style to apply

            // Save the workbook
            workbook.Save("PivotHeaderFormatted.xlsx");
        }
    }
}
