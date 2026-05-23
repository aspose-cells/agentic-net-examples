using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvUtf32Example
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");

            // Populate a large dataset (example: 10,000 rows)
            for (int i = 2; i <= 10001; i++)
            {
                sheet.Cells[i - 1, 0].PutValue($"User{i - 1}");
                sheet.Cells[i - 1, 1].PutValue(20 + (i % 30));
            }

            // Create CSV save options and set UTF‑32 encoding
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.Encoding = Encoding.UTF32;   // Specify UTF‑32 encoding
            csvOptions.Separator = ',';            // Use comma as separator

            // Save the workbook as CSV with the custom options
            workbook.Save("LargeDataset.csv", csvOptions);
        }
    }
}