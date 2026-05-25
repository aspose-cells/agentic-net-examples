using System;
using Aspose.Cells;

namespace AsposeCellsReportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate product name
            sheet.Cells["A2"].PutValue("Laptop");

            // Populate price value
            sheet.Cells["B2"].PutValue(999.99);

            // Apply built‑in currency format (Number = 4) to the price cell
            Style priceStyle = sheet.Cells["B2"].GetStyle();
            priceStyle.Number = 4; // Currency format
            sheet.Cells["B2"].SetStyle(priceStyle);

            // Retrieve the formatted price using GetStringValue with DisplayString strategy
            // This includes the currency symbol as it appears in Excel
            string formattedPrice = sheet.Cells["B2"]
                .GetStringValue(CellValueFormatStrategy.DisplayString);

            // Build a user‑friendly report line
            string reportLine = $"Product: {sheet.Cells["A2"].StringValue}, Price: {formattedPrice}";

            // Output the report line
            Console.WriteLine(reportLine);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ProductReport.xlsx");
        }
    }
}