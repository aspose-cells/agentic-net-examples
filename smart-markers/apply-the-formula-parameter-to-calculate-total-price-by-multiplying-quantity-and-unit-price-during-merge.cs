// Title: Calculate total price with a multiplication formula and create a centered merged title row using Aspose.Cells for .NET (C#)
// Common Searches: how to set a multiplication formula for total price in Aspose.Cells C# example | Aspose.Cells C# merge first row cells and center title | using Aspose.Cells to calculate quantity times unit price in Excel | C# code to apply a formula to a range of cells with Aspose.Cells | Aspose.Cells tutorial for merging cells and applying formulas together
// Tags: multiplication formula total price Aspose.Cells | merge cells for title row Aspose.Cells | center align merged title Aspose.Cells | set cell formula C# Aspose.Cells | calculate quantity unit price Excel Aspose.Cells | apply formulas during cell merge Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, adds Quantity, Unit Price, and Total Price columns, populates sample data, assigns a multiplication formula to each Total Price cell, merges the first row to form a centered title, and saves the file as TotalPriceCalculation.xlsx using Aspose.Cells for .NET.
class CalculateTotalPrice
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers
            sheet.Cells["A1"].PutValue("Quantity");
            sheet.Cells["B1"].PutValue("Unit Price");
            sheet.Cells["C1"].PutValue("Total Price");

            // Sample data
            sheet.Cells["A2"].PutValue(5);
            sheet.Cells["B2"].PutValue(12.5);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["B3"].PutValue(7.8);
            sheet.Cells["A4"].PutValue(10);
            sheet.Cells["B4"].PutValue(4.2);

            // Apply formula to calculate total price (Quantity * Unit Price)
            for (int row = 2; row <= 4; row++)
            {
                // C{row} = A{row} * B{row}
                string formula = $"=A{row}*B{row}";
                sheet.Cells[$"C{row}"].Formula = formula;
            }

            // Merge cells for a title spanning all three columns
            sheet.Cells.Merge(0, 0, 1, 3); // Merge A1:C1
            sheet.Cells["A1"].PutValue("Order Summary");

            // Center align the title using the cell's style
            Style titleStyle = sheet.Cells["A1"].GetStyle();
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            sheet.Cells["A1"].SetStyle(titleStyle);

            // Save the workbook
            workbook.Save("TotalPriceCalculation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
