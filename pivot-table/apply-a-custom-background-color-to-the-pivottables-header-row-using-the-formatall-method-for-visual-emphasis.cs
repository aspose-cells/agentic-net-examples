// Title: C# – Apply a Custom Background Color to a PivotTable Header Row with PivotTable.FormatAll (Aspose.Cells)
// Description: This example creates a workbook, adds sample data, builds a PivotTable, defines a Style with a LightBlue solid fill and bold font, and applies the style to the entire PivotTable using PivotTable.FormatAll. The call colors the header row (and other cells) before saving the file as XLSX.
// Keywords: Aspose.Cells | PivotTable | FormatAll | C# | .NET | header background color | custom style | solid fill | bold font | Excel export | sample code
// Common Searches: Aspose.Cells change PivotTable header background color C# | PivotTable.FormatAll example .NET | How to style PivotTable header row with Aspose.Cells | Apply solid fill to PivotTable header using C# | C# code to format entire PivotTable with a style
// Developer Intent: The developer wants to highlight the PivotTable header row by applying a custom background color via the FormatAll method.
// Use Cases: Add a light‑blue background to the header row of a sales summary PivotTable for clearer visual separation. | Create a uniform report style by applying a bold font and solid fill to all cells of multiple PivotTables in an automated .NET workflow. | Reuse a single Style object to format several PivotTables across worksheets with one FormatAll call, reducing code duplication.
// AI Prompts: Write C# code that defines a dark‑gray Style and applies it only to the header row of a PivotTable using Aspose.Cells. | Show how to use conditional formatting together with FormatAll to give the header row a blue background and data rows a white background. | Demonstrate reusing a Style object to format three different PivotTables in the same workbook with Aspose.Cells .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds sample data, builds a PivotTable, defines a Style with a LightBlue solid fill and bold font, and applies the style to the entire PivotTable using PivotTable.FormatAll. The call colors the header row (and other cells) before saving the file as XLSX.
    public class PivotTableHeaderRowFormatAllDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].Value = "Category";
                worksheet.Cells["B1"].Value = "Amount";
                worksheet.Cells["A2"].Value = "Fruit";
                worksheet.Cells["B2"].Value = 120;
                worksheet.Cells["A3"].Value = "Vegetable";
                worksheet.Cells["B3"].Value = 80;
                worksheet.Cells["A4"].Value = "Fruit";
                worksheet.Cells["B4"].Value = 150;
                worksheet.Cells["A5"].Value = "Vegetable";
                worksheet.Cells["B5"].Value = 60;

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Category, data = Sum of Amount
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Calculate the pivot table data
                pivotTable.CalculateData();

                // Create a style with a custom background color for emphasis
                Style headerStyle = workbook.CreateStyle();
                headerStyle.ForegroundColor = Color.LightBlue;   // Desired background color
                headerStyle.Pattern = BackgroundType.Solid;     // Apply solid fill
                headerStyle.Font.IsBold = true;                 // Optional: make header text bold

                // Apply the style to the entire pivot table using FormatAll.
                // This will color the header row (and all other cells) with the specified background.
                pivotTable.FormatAll(headerStyle);

                // Save the workbook to a file
                workbook.Save("PivotTableHeaderRowFormatAllDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
