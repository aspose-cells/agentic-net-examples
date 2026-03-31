using System;
using Aspose.Cells;

namespace AsposeCellsSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Save the workbook as XLSX using default settings
            // The Save method overload (string, SaveFormat) is used as per the provided rules
            workbook.Save("SampleOutput.xlsx", SaveFormat.Xlsx);

            // Optionally, demonstrate saving to other formats with default settings
            workbook.Save("SampleOutput.xls", SaveFormat.Excel97To2003);
            workbook.Save("SampleOutput.csv", SaveFormat.Csv);
            workbook.Save("SampleOutput.pdf", SaveFormat.Pdf);
            workbook.Save("SampleOutput.html", SaveFormat.Html);

            Console.WriteLine("Workbook saved in multiple formats.");
        }
    }
}