// Title: Aspose.Cells for .NET: Change Pivot Table Header Font Color to Black (C#)
// Description: C# example that creates a workbook, builds a pivot table, calculates it, defines a Style with a black font, and applies the style to the pivot table header cell using PivotTable.Format before saving the file.
// Keywords: Aspose.Cells | C# | .NET | pivot table header font color | black font style | PivotTable.Format | Excel styling with Aspose | code example | format pivot header
// Common Searches: Aspose.Cells set pivot header font color | C# change pivot table header text to black | format pivot table header cell Aspose | apply style to pivot field header .NET | how to color pivot table header in Aspose.Cells
// Developer Intent: Apply a black font color to a pivot table header cell using Aspose.Cells for .NET.
// Use Cases: Create a consistent black‑font style for all pivot table headers in a financial report. | Programmatically highlight specific pivot headers after generating a dynamic Excel workbook. | Reuse a single Style object to format multiple pivot table headers across several worksheets.
// AI Prompts: Show me C# code that changes a pivot table header font color to black with Aspose.Cells. | How can I apply a black font style to a specific pivot table header cell using PivotTable.Format? | Explain how to format multiple pivot table header cells with the same black font style in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, builds a pivot table, calculates it, defines a Style with a black font, and applies the style to the pivot table header cell using PivotTable.Format before saving the file.
    public class ChangePivotHeaderFontColor
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(40);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Value as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Value

            // Calculate the pivot table so that header cells are generated
            pivotTable.CalculateData();

            // Create a style that sets the font color to black
            Style blackFontStyle = workbook.CreateStyle();
            blackFontStyle.Font.Color = Color.Black;

            // Apply the style to the header cell of the row field
            pivotTable.Format(0, 0, blackFontStyle);

            // Save the workbook with the applied formatting
            workbook.Save("PivotHeaderBlackFont.xlsx");
        }
    }
}
