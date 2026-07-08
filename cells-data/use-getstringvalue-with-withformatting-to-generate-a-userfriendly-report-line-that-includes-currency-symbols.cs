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
            Cells cells = sheet.Cells;

            // Put a numeric value that represents an amount
            Cell amountCell = cells["B2"];
            amountCell.PutValue(1234.56);

            // Apply a built‑in currency format (Number format index 4)
            Style currencyStyle = amountCell.GetStyle();
            currencyStyle.Number = 4; // Currency format
            amountCell.SetStyle(currencyStyle);

            // Retrieve the formatted string using GetStringValue with DisplayString strategy
            string formattedAmount = amountCell.GetStringValue(CellValueFormatStrategy.DisplayString);

            // Build a user‑friendly report line that includes the currency symbol
            string reportLine = $"Total Sales: {formattedAmount}";

            // Output the report line to the console
            Console.WriteLine(reportLine);

            // Optionally, write the report line into another cell for demonstration
            cells["A4"].PutValue(reportLine);

            // Save the workbook (the file will contain the value and the report line)
            workbook.Save("ReportDemo.xlsx");
        }
    }
}