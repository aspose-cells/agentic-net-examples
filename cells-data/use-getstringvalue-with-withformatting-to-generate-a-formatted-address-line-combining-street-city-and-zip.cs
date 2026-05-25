using System;
using Aspose.Cells;

namespace AsposeCellsAddressExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with address components
            sheet.Cells["A1"].PutValue("123 Main St");   // Street
            sheet.Cells["B1"].PutValue("Springfield");  // City
            sheet.Cells["C1"].PutValue("12345");        // ZIP code

            // Create a formula that concatenates the components into a single address line
            // Result will look like: 123 Main St, Springfield 12345
            sheet.Cells["D1"].Formula = "=A1 & \", \" & B1 & \" \" & C1";

            // Retrieve the formatted string value of the formula cell using DisplayString strategy
            string formattedAddress = sheet.Cells["D1"].GetStringValue(CellValueFormatStrategy.DisplayString);

            // Output the formatted address
            Console.WriteLine("Formatted Address: " + formattedAddress);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("FormattedAddress.xlsx");
        }
    }
}