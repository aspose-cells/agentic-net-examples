using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaWithNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["B3"].PutValue(20);

            // -------------------------------------------------
            // Define a named range that covers the desired print area
            // -------------------------------------------------
            // Add a new name to the workbook's name collection
            int nameIndex = workbook.Worksheets.Names.Add("MyPrintArea");
            // Set the reference of the name to the actual cell range (absolute reference)
            // Note: The reference must start with an equal sign.
            workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$B$3";

            // -------------------------------------------------
            // Assign the named range to the worksheet's print area
            // -------------------------------------------------
            // The PrintArea property can accept a named range identifier.
            sheet.PageSetup.PrintArea = "MyPrintArea";

            // -------------------------------------------------
            // Save the workbook (the print area will be based on the named range)
            // -------------------------------------------------
            workbook.Save("PrintAreaWithNamedRange.xlsx");

            Console.WriteLine("Workbook saved with print area set via named range.");
        }
    }
}