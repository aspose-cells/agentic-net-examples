using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set the default column width (in characters) before adding any data
        // Author note: establishing a base width ensures consistent column sizing
        cells.StandardWidth = 20;

        // Populate sample data
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");
        cells["A2"].PutValue("Alice");
        cells["B2"].PutValue(30);
        cells["A3"].PutValue("Bob");
        cells["B3"].PutValue(25);

        // Save the workbook
        workbook.Save("StandardWidthDemo.xlsx");
    }
}