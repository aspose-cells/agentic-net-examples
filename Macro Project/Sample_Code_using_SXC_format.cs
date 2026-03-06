using Aspose.Cells;
using System;

namespace AsposeCellsSxcDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Save the workbook in StarOffice Calc (.sxc) format
            workbook.Save("Sample.sxc", SaveFormat.Sxc);

            // Load the saved .sxc file to verify it was saved correctly
            Workbook loadedWorkbook = new Workbook("Sample.sxc");

            // Output a cell value from the loaded workbook
            Console.WriteLine("Loaded cell A2 value: " + loadedWorkbook.Worksheets[0].Cells["A2"].StringValue);
        }
    }
}