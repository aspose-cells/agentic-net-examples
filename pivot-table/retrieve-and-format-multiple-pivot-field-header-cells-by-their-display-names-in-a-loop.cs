// Title: Aspose.Cells for .NET – Retrieve PivotTable header cells by display name in a loop and apply a custom style (C#)
// Description: This C# sample builds a workbook, inserts sample data, creates a PivotTable, defines a bold light‑gray style, iterates through row, column and data fields, fetches each header cell with GetCellByDisplayName, applies the style, and saves the file as PivotHeaderFormatting.xlsx.
// Keywords: Aspose.Cells C# pivot table header formatting | GetCellByDisplayName | apply style to pivot headers | loop through pivot fields | Aspose.Cells PivotTable example | format pivot table header cells programmatically | C# Excel pivot table styling | Aspose.Cells API GetCellByDisplayName
// Common Searches: How to format pivot table header cells using Aspose.Cells C# | Retrieve pivot field header by display name Aspose.Cells | Loop over PivotTable fields to set style .NET | Apply background color to Excel pivot headers with Aspose | C# code for styling row and column headers in a PivotTable
// Developer Intent: Automatically style all PivotTable header cells by iterating over their display names.
// Use Cases: Standardize appearance of row field headers (e.g., Category) across generated reports. | Apply consistent visual cues to column field headers (e.g., SubCategory) for improved readability. | Highlight data field headers (e.g., Amount) with bold font and gray background in Excel outputs. | Integrate header styling into automated Excel report generation pipelines.
// AI Prompts: Generate C# code that uses Aspose.Cells to loop through PivotTable row, column, and data fields, retrieve each header cell by display name, and apply a bold light‑gray style. | Show how to create a reusable method in Aspose.Cells for .NET that formats multiple pivot field headers based on their display names with proper error handling. | Provide an example of applying a solid background color and bold font to all PivotTable headers using GetCellByDisplayName in a C# workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHeaderFormatting
{
    // This C# sample builds a workbook, inserts sample data, creates a PivotTable, defines a bold light‑gray style, iterates through row, column and data fields, fetches each header cell with GetCellByDisplayName, applies the style, and saves the file as PivotHeaderFormatting.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "SubCategory";
                sheet.Cells["C1"].Value = "Amount";

                sheet.Cells["A2"].Value = "Food";
                sheet.Cells["B2"].Value = "Fruit";
                sheet.Cells["C2"].Value = 120;

                sheet.Cells["A3"].Value = "Food";
                sheet.Cells["B3"].Value = "Vegetable";
                sheet.Cells["C3"].Value = 80;

                sheet.Cells["A4"].Value = "Beverage";
                sheet.Cells["B4"].Value = "Tea";
                sheet.Cells["C4"].Value = 50;

                sheet.Cells["A5"].Value = "Beverage";
                sheet.Cells["B5"].Value = "Coffee";
                sheet.Cells["C5"].Value = 70;

                // Add a pivot table
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // SubCategory as column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Amount as data field

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Define a style to apply to header cells
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;

                // Helper method to apply style to a cell if it exists
                void ApplyStyleToHeader(string displayName)
                {
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    if (headerCell != null)
                    {
                        headerCell.SetStyle(headerStyle);
                    }
                }

                // Apply style to row field headers
                for (int i = 0; i < pivotTable.RowFields.Count; i++)
                {
                    string displayName = pivotTable.RowFields[i].DisplayName;
                    ApplyStyleToHeader(displayName);
                }

                // Apply style to column field headers
                for (int i = 0; i < pivotTable.ColumnFields.Count; i++)
                {
                    string displayName = pivotTable.ColumnFields[i].DisplayName;
                    ApplyStyleToHeader(displayName);
                }

                // Apply style to data field headers
                for (int i = 0; i < pivotTable.DataFields.Count; i++)
                {
                    string displayName = pivotTable.DataFields[i].DisplayName;
                    ApplyStyleToHeader(displayName);
                }

                // Save the workbook
                workbook.Save("PivotHeaderFormatting.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
