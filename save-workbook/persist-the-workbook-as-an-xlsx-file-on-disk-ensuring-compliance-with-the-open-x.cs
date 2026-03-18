using System;
using Aspose.Cells;

namespace AsposeCellsSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Optional: enforce strict OOXML compliance
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Persist the workbook as an XLSX file on disk (save rule)
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);

            // Clean up resources
            workbook.Dispose();

            Console.WriteLine("Workbook saved successfully as Result.xlsx");
        }
    }
}