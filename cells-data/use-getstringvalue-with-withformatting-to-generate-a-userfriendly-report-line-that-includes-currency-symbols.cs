// Title: Create a currency‑formatted report line in C# using Aspose.Cells GetStringValue with DisplayString
// AI Prompts: Write C# code that sets a cell to the built‑in currency number format, calls GetStringValue with CellValueFormatStrategy.DisplayString, and concatenates the result into a readable report string. | Show how to retrieve a cell's formatted value including the currency symbol using Aspose.Cells and embed it with another cell's text in a single output line. | Demonstrate using GetStringValue(DisplayString) to obtain a locale‑aware currency string and combine it with product information for console output.
// Common Searches: Aspose.Cells C# get cell value with currency symbol using GetStringValue | How to format price as currency and read it back as string in Aspose.Cells .NET | C# Aspose.Cells DisplayString strategy for formatted monetary values | Create a report line with product name and formatted price in Aspose.Cells workbook | Retrieve formatted cell value including currency sign in Aspose.Cells API
// Tags: GetStringValue DisplayString currency formatting | apply built-in currency number format Aspose.Cells | concatenate cell StringValue with formatted price C# | Aspose.Cells formatted monetary value retrieval | report line generation using Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsReportDemo
{
    // // Applies the built‑in currency number format to a cell, retrieves the formatted value with GetStringValue using the DisplayString strategy, builds a console report line that combines the product name and the currency‑formatted price, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put product name in A1
            Cell productCell = cells["A1"];
            productCell.PutValue("Laptop");

            // Put price value in B1
            Cell priceCell = cells["B1"];
            priceCell.PutValue(999.99);

            // Apply built‑in currency format (Number = 4) to the price cell
            Style priceStyle = priceCell.GetStyle();
            priceStyle.Number = 4; // Currency format
            priceCell.SetStyle(priceStyle);

            // Retrieve the formatted price string using GetStringValue with DisplayString strategy
            string formattedPrice = priceCell.GetStringValue(CellValueFormatStrategy.DisplayString);

            // Build a user‑friendly report line that includes the currency symbol
            string reportLine = $"Product: {productCell.StringValue}, Price: {formattedPrice}";

            // Output the report line
            Console.WriteLine(reportLine);

            // Optionally save the workbook (uses the standard create/save lifecycle)
            workbook.Save("ReportDemo.xlsx");
        }
    }
}
