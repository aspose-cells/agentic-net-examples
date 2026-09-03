// Title: Generate an Excel workbook with twelve month‑named worksheets and identical sample tables using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to create a new workbook, clear the default sheet, add worksheets named January through December, insert a header row (Item, Quantity, Price) and three sample data rows into each sheet, then export the file as an .xlsx workbook. | Demonstrate how to iterate over an array of month names in C# with Aspose.Cells to programmatically add and name worksheets, fill each with the same data table, and save the workbook.
// Common Searches: Aspose.Cells C# create workbook with worksheets for each month | add multiple sheets named January February Aspose.Cells .NET | populate identical tables across many worksheets using Aspose.Cells | remove default sheet and add custom worksheets in Aspose.Cells C# | save Excel file with month tabs using Aspose.Cells for .NET
// Tags: add month worksheets Aspose.Cells | write header and sample rows Aspose.Cells | clear initial worksheet Aspose.Cells | export to .xlsx using Aspose.Cells | duplicate sample data across sheets Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, clears the default sheet, adds twelve worksheets named after each month, writes a header and three rows of sample item data to every sheet, and saves the result as MonthsWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Remove the default worksheet that comes with a new workbook
        workbook.Worksheets.Clear();

        // Array of month names to be used as worksheet names
        string[] months = new string[]
        {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };

        // Sample data to populate each worksheet
        string[] items = { "Item A", "Item B", "Item C" };
        int[] quantities = { 10, 20, 30 };
        double[] prices = { 1.5, 2.0, 3.5 };

        // Loop through each month, create a worksheet, name it, and fill with sample data
        for (int i = 0; i < months.Length; i++)
        {
            // Add a new worksheet and get its reference
            Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet.Name = months[i];

            // Write header row
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            // Populate sample rows starting from row 2
            for (int row = 0; row < items.Length; row++)
            {
                int excelRow = row + 2; // Excel rows are 1-indexed
                sheet.Cells[excelRow - 1, 0].PutValue(items[row]);      // Column A
                sheet.Cells[excelRow - 1, 1].PutValue(quantities[row]); // Column B
                sheet.Cells[excelRow - 1, 2].PutValue(prices[row]);     // Column C
            }
        }

        // Save the workbook to a file
        workbook.Save("MonthsWorkbook.xlsx");
    }
}
