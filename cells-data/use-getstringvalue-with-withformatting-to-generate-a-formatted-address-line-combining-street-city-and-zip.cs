// Title: Create a formatted address string from separate cells using GetStringValue with DisplayString in Aspose.Cells for .NET
// AI Prompts: Read cells A1, B1, and C1 with GetStringValue using CellValueFormatStrategy.DisplayString and concatenate them into a single string "street, city zip" in C#. | Write a C# program that extracts street, city, and zip from a worksheet, builds a formatted address line, places it in D1, and saves the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# read cell with DisplayString and concatenate values | How to format address from multiple cells using GetStringValue in Aspose.Cells | Combine street, city, zip into one cell Aspose.Cells .NET example | GetStringValue CellValueFormatStrategy.DisplayString usage for Excel address formatting
// Tags: GetStringValue DisplayString formatting Aspose.Cells | concatenate address components Excel .NET | write formatted address to worksheet cell | Aspose.Cells cell value retrieval and formatting | save workbook with formatted address Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsAddressDemo
{
    // The example creates a workbook, fills A1‑C1 with street, city, and zip, reads each cell using GetStringValue with CellValueFormatStrategy.DisplayString, builds a formatted address string "street, city zip", writes it to D1, and saves the file as FormattedAddress.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with address components
            cells["A1"].PutValue("123 Main St");   // Street
            cells["B1"].PutValue("Springfield");   // City
            cells["C1"].PutValue("98765");         // Zip

            // Retrieve each component using GetStringValue with formatting (DisplayString)
            string street = cells["A1"].GetStringValue(CellValueFormatStrategy.DisplayString);
            string city   = cells["B1"].GetStringValue(CellValueFormatStrategy.DisplayString);
            string zip    = cells["C1"].GetStringValue(CellValueFormatStrategy.DisplayString);

            // Combine components into a formatted address line
            string formattedAddress = $"{street}, {city} {zip}";

            // Write the formatted address back to the worksheet (optional)
            cells["D1"].PutValue(formattedAddress);

            // Save the workbook
            workbook.Save("FormattedAddress.xlsx");
        }
    }
}
