using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate address components
        cells["A1"].PutValue("123 Main St");   // Street
        cells["A2"].PutValue("Springfield");   // City
        cells["A3"].PutValue("12345");         // ZIP code

        // Create a formula that concatenates the components with proper separators
        cells["A4"].Formula = "=A1 & \", \" & A2 & \" \" & A3";

        // Evaluate the formula so the cell contains the resulting value
        workbook.CalculateFormula();

        // Retrieve the formatted address line using GetStringValue with DisplayString strategy
        string formattedAddress = cells["A4"].GetStringValue(CellValueFormatStrategy.DisplayString);

        // Display the result
        Console.WriteLine("Formatted Address: " + formattedAddress);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("FormattedAddress.xlsx");
    }
}