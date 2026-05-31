using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);

            // Configure text save options to use semicolon as the column delimiter
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ';',          // Set semicolon delimiter
                Encoding = Encoding.UTF8 // Optional: set encoding
            };

            // Save the workbook as CSV using the custom options
            workbook.Save("output.csv", saveOptions);

            Console.WriteLine("Workbook saved as CSV with semicolon delimiter.");
        }
    }
}