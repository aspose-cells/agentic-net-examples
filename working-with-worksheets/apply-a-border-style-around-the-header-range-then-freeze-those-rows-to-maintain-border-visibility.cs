// Title: Add a Thick Red Outline Border to a Header Row and Freeze It with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, writes a header in A1:D1, applies a thick red outline border using SetOutlineBorders, freezes the header row with FreezePanes at A2, and saves the file as HeaderBorderAndFreeze.xlsx.
// Keywords: Aspose.Cells | C# | SetOutlineBorders | FreezePanes | header row border | Excel styling | outline border thick red | freeze top row | Excel workbook generation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set outline border C# | Freeze first row Aspose.Cells .NET | Apply red border to header Excel using Aspose | How to freeze panes after styling header Aspose.Cells | C# code for border and freeze panes in Excel
// Developer Intent: Apply a thick red outline border to the header range (A1:D1) and freeze that row so it remains visible while scrolling.
// Use Cases: Sales reports where column titles are highlighted with a red border and stay fixed during scrolling. | Invoice templates that keep the header row visible and visually distinct for easy reference. | Large data exports where a styled, frozen header improves readability and analysis.
// AI Prompts: Write C# code using Aspose.Cells to apply a double blue border to header range A1:F1 and freeze the top two rows. | Show how to set different border styles for multiple header ranges and freeze both rows and columns in an Aspose.Cells workbook. | Provide an example that adds a thick green outline border to a header row, freezes the pane at cell B2, and saves the workbook as an .xlsx file.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHeaderBorderAndFreeze
{
    // C# example that creates a workbook, writes a header in A1:D1, applies a thick red outline border using SetOutlineBorders, freezes the header row with FreezePanes at A2, and saves the file as HeaderBorderAndFreeze.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data with a header row (A1:D1)
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Category");
                cells["C1"].PutValue("Price");
                cells["D1"].PutValue("Quantity");

                // Fill a few data rows for visibility
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue("Fruit");
                cells["C2"].PutValue(1.20);
                cells["D2"].PutValue(50);

                cells["A3"].PutValue("Carrot");
                cells["B3"].PutValue("Vegetable");
                cells["C3"].PutValue(0.80);
                cells["D3"].PutValue(30);

                // Define the header range (first row)
                Aspose.Cells.Range headerRange = cells.CreateRange("A1:D1");

                // Apply a thick red outline border around the header range
                headerRange.SetOutlineBorders(CellBorderType.Thick, Color.Red);

                // Freeze the header row so it stays visible while scrolling
                // Freeze at cell A2 (row index 2) with 1 frozen row and 0 frozen columns
                worksheet.FreezePanes("A2", 1, 0);

                // Save the workbook
                string outputPath = "HeaderBorderAndFreeze.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
