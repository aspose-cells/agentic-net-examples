using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Row{i - 1} Col1");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Define a named range called "ReportArea" that covers the data region
            // Add the name to the workbook's collection of names
            int nameIndex = workbook.Worksheets.Names.Add("ReportArea");
            // Set the reference of the named range (absolute address)
            workbook.Worksheets.Names[nameIndex].RefersTo = $"=Sheet1!$A$1:$B$10";

            // Set the worksheet's print area to the named range.
            // Aspose.Cells allows assigning the name directly to PrintArea.
            sheet.PageSetup.PrintArea = "ReportArea";

            // Save the workbook (lifecycle: save)
            workbook.Save("ReportAreaDemo.xlsx");

            Console.WriteLine("Workbook saved with print area set to named range 'ReportArea'.");
        }
    }
}