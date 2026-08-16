// Title: Aspose.Cells .NET: Apply a Dark‑Green Font Style to a Pivot Table Using PivotTable.Format
// Description: This C# example demonstrates how to create a Workbook, add sample data, build a pivot table, define a Style with a dark‑green bold font, and apply that style to the pivot table's data body range via the PivotTable.Format method before saving the file.
// Keywords: Aspose.Cells pivot table styling | C# PivotTable.Format example | set font color Aspose.Cells | custom workbook style .NET | format pivot table cells
// Common Searches: how to change pivot table font color with Aspose.Cells | apply style to PivotTable.DataBodyRange C# | Aspose.Cells format specific area of pivot table | create bold green font style for pivot table
// Developer Intent: Generate a Style object, configure its font color, and apply it to a selected area of a pivot table using PivotTable.Format.
// Use Cases: Emphasize high‑value sales rows with a green bold font for quick visual scanning. | Maintain consistent formatting across multiple reports by reusing a single style definition. | Programmatically adjust cell appearance based on business thresholds, such as coloring overdue items.
// AI Prompts: Write C# code that creates a red italic Style and applies it to the column headers of a pivot table with Aspose.Cells. | Show how to use PivotTable.Format to give different background colors to data rows and total rows in a pivot table. | Provide an example that sets both font size and background shade for a custom CellArea in a pivot table using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotStyleDemo
{
    // This C# example demonstrates how to create a Workbook, add sample data, build a pivot table, define a Style with a dark‑green bold font, and apply that style to the pivot table's data body range via the PivotTable.Format method before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].Value = "Product";
            worksheet.Cells["B1"].Value = "Region";
            worksheet.Cells["C1"].Value = "Sales";

            worksheet.Cells["A2"].Value = "Laptop";
            worksheet.Cells["B2"].Value = "North";
            worksheet.Cells["C2"].Value = 1200;

            worksheet.Cells["A3"].Value = "Desktop";
            worksheet.Cells["B3"].Value = "South";
            worksheet.Cells["C3"].Value = 800;

            worksheet.Cells["A4"].Value = "Tablet";
            worksheet.Cells["B4"].Value = "East";
            worksheet.Cells["C4"].Value = 500;

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Ensure the pivot table is calculated
            pivotTable.CalculateData();

            // Create a custom style and set its font color
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Color = Color.DarkGreen;   // Set desired font color
            customStyle.Font.IsBold = true;             // Optional: make the font bold

            // Define the area to format – here we format the entire data body range
            CellArea dataArea = pivotTable.DataBodyRange;

            // Apply the custom style to the specified area using PivotTable.Format
            pivotTable.Format(dataArea, customStyle);

            // Save the workbook to a file
            workbook.Save("PivotTableCustomFontColor.xlsx", SaveFormat.Xlsx);
        }
    }
}
