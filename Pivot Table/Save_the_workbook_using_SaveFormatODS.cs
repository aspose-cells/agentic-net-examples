using System;
using Aspose.Cells;

namespace AsposeCellsOdsSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B2"].PutValue(12345);

            // Save the workbook in ODS format using the Save(string, SaveFormat) overload
            workbook.Save("SampleOutput.ods", SaveFormat.Ods);

            Console.WriteLine("Workbook saved as ODS successfully.");
        }
    }
}