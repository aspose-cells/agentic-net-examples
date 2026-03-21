using System;
using Aspose.Cells;

namespace CombineMultipleWorkbooksDemo
{
    class Program
    {
        static void Main()
        {
            // Create first source workbook and add some data
            Workbook sourceWorkbook1 = new Workbook();
            Worksheet sheet1 = sourceWorkbook1.Worksheets[0];
            sheet1.Name = "Source1";
            sheet1.Cells["A1"].PutValue("Data from Workbook 1");
            sheet1.Cells["B2"].PutValue(123);

            // Create second source workbook and add some data
            Workbook sourceWorkbook2 = new Workbook();
            Worksheet sheet2 = sourceWorkbook2.Worksheets[0];
            sheet2.Name = "Source2";
            sheet2.Cells["A1"].PutValue("Data from Workbook 2");
            sheet2.Cells["C3"].PutValue(456.78);

            // Create third source workbook and add some data
            Workbook sourceWorkbook3 = new Workbook();
            Worksheet sheet3 = sourceWorkbook3.Worksheets[0];
            sheet3.Name = "Source3";
            sheet3.Cells["A1"].PutValue("Data from Workbook 3");
            sheet3.Cells["D4"].PutValue(DateTime.Now);

            // Create destination workbook (empty)
            Workbook destWorkbook = new Workbook();

            // Combine the source workbooks into the destination workbook
            destWorkbook.Combine(sourceWorkbook1);
            destWorkbook.Combine(sourceWorkbook2);
            destWorkbook.Combine(sourceWorkbook3);

            // Save the combined workbook to disk
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbooks combined and saved as 'CombinedWorkbook.xlsx'.");
        }
    }
}