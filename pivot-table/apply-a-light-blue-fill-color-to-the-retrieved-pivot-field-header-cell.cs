// Title: Aspose.Cells for .NET – Set Light Blue Fill on a Pivot Table Header Cell (C#)
// Description: C# example that creates a workbook, builds a pivot table, locates the row‑field header cell ("Product") with the Find method, defines a solid LightBlue style, enables cell shading, applies the style to the header, and saves the file as PivotHeaderLightBlue.xlsx.
// Keywords: Aspose.Cells | C# | pivot table header background | light blue fill | cell style Aspose.Cells | Excel pivot formatting | StyleFlag | Find method | sample code | GitHub example
// Common Searches: Aspose.Cells change pivot header color C# | set background fill for pivot table header Aspose.Cells | apply light blue style to pivot field header .NET | format pivot table header cell programmatically | C# example for styling pivot table headers
// Developer Intent: Apply a solid LightBlue background to a specific pivot table header cell.
// Use Cases: Highlight row‑field headers to improve visual hierarchy in automated Excel reports. | Enforce corporate color schemes on pivot tables before distribution. | Batch‑process workbooks and consistently style pivot headers across multiple files.
// AI Prompts: Generate C# code with Aspose.Cells that colors a pivot table header cell based on its text. | Show how to style all column headers of a pivot table with a gradient fill using Aspose.Cells for .NET. | Explain alternative ways to locate pivot header cells without using Find, such as by index or field position.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotHeaderStyle
{
    // C# example that creates a workbook, builds a pivot table, locates the row‑field header cell ("Product") with the Find method, defines a solid LightBlue style, enables cell shading, applies the style to the header, and saves the file as PivotHeaderLightBlue.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Bike";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Bike";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 1500;

            sheet.Cells["A4"].Value = "Car";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 2000;

            sheet.Cells["A5"].Value = "Car";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 2500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Region as column field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table so that header cells are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the header cell for the row field ("Product")
            // After calculation the header text appears in the pivot table; we locate it using Find
            Cell headerCell = sheet.Cells.Find("Product", null);
            if (headerCell != null)
            {
                // Create a style with LightBlue fill
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightBlue;

                // Enable cell shading so the fill color is applied
                StyleFlag flag = new StyleFlag();
                flag.CellShading = true;

                // Apply the style to the found header cell
                headerCell.SetStyle(style, flag);
            }

            // Save the workbook
            workbook.Save("PivotHeaderLightBlue.xlsx");
        }
    }
}
