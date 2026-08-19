// Title: Aspose.Cells .NET: Use GetStringValue with DisplayString to build a formatted address line
// Description: This example shows how to create a workbook, store street, city, and ZIP values in cells A1‑C1, concatenate them with a formula in D1, calculate the sheet, and retrieve the full address as a formatted string using GetStringValue with the CellValueFormatStrategy.DisplayString strategy. The result is printed and the workbook saved as FormattedAddress.xlsx.
// Keywords: Aspose.Cells GetStringValue | CellValueFormatStrategy.DisplayString | C# address concatenation | Excel formatted string extraction | .NET Aspose.Cells example | formatted address line
// Common Searches: Aspose.Cells GetStringValue formatting example | How to retrieve formatted cell text in C# | Concatenate address cells in Aspose.Cells | DisplayString strategy Aspose.Cells | GetStringValue address line .NET
// Developer Intent: Extract a single, human‑readable address string from a worksheet cell by applying GetStringValue with a display‑formatting strategy.
// Use Cases: Generate mailing‑label strings directly from spreadsheet data. | Provide UI components or reports with a ready‑to‑display address without extra string handling. | Create API payloads that require a full address line extracted from Excel files.
// AI Prompts: Extend the code to include a state column and format the address as "Street, City, State ZIP" using GetStringValue. | Show how to use GetStringValue with CellValueFormatStrategy.DisplayString to format numeric cells as currency. | Provide a robust method to skip empty address parts when concatenating cells and retrieving the formatted string.

using System;
using Aspose.Cells;

namespace AsposeCellsAddressExample
{
    // This example shows how to create a workbook, store street, city, and ZIP values in cells A1‑C1, concatenate them with a formula in D1, calculate the sheet, and retrieve the full address as a formatted string using GetStringValue with the CellValueFormatStrategy.DisplayString strategy. The result is printed and the workbook saved as FormattedAddress.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate individual address components
            cells["A1"].PutValue("123 Main St");   // Street
            cells["B1"].PutValue("Springfield");   // City
            cells["C1"].PutValue("98765");         // ZIP

            // Create a formula that concatenates the components into a full address line
            // Example result: "123 Main St, Springfield 98765"
            cells["D1"].Formula = "A1 & \", \" & B1 & \" \" & C1";

            // Calculate the formula so the cell contains the resulting value
            workbook.CalculateFormula();

            // Retrieve the formatted address line using GetStringValue with DisplayString strategy
            string formattedAddress = cells["D1"].GetStringValue(CellValueFormatStrategy.DisplayString);

            // Output the result
            Console.WriteLine("Formatted Address: " + formattedAddress);

            // Save the workbook (lifecycle: save)
            workbook.Save("FormattedAddress.xlsx");
        }
    }
}
