using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvUtf16Example
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Configure CSV save options with UTF-16 (Unicode) encoding
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.Unicode // UTF-16 encoding
            };

            // Save the workbook as CSV using the save options (lifecycle rule: save)
            workbook.Save("output_utf16.csv", saveOptions);

            Console.WriteLine("Workbook exported to CSV with UTF-16 encoding.");
        }
    }
}