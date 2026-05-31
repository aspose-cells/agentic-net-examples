using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    class Program
    {
        static void Main()
        {
            // Create a sample workbook with multiple worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Customers";
            workbook.Worksheets[0].Cells["A1"].PutValue("Name");
            workbook.Worksheets[0].Cells["B1"].PutValue("Country");
            workbook.Worksheets[0].Cells["A2"].PutValue("Alice");
            workbook.Worksheets[0].Cells["B2"].PutValue("USA");

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Orders");
            sheet2.Cells["A1"].PutValue("OrderId");
            sheet2.Cells["B1"].PutValue("Amount");
            sheet2.Cells["A2"].PutValue(1001);
            sheet2.Cells["B2"].PutValue(250.75);

            // Define a custom delimiter (e.g., semicolon) and UTF‑8 encoding
            char customDelimiter = ';';
            Encoding utf8 = Encoding.UTF8;

            // Export each worksheet to its own CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as active so that ExportAllSheets = false works correctly
                workbook.Worksheets.ActiveSheetIndex = i;

                // Configure text save options
                TxtSaveOptions saveOptions = new TxtSaveOptions();
                saveOptions.Separator = customDelimiter;   // custom delimiter
                saveOptions.Encoding = utf8;               // UTF‑8 encoding
                saveOptions.ExportAllSheets = false;       // export only the active sheet

                // Build output file name using the worksheet name
                string outputFile = $"{workbook.Worksheets[i].Name}.csv";

                // Save the active worksheet as CSV with the specified options
                workbook.Save(outputFile, saveOptions);
            }

            Console.WriteLine("CSV files have been generated for each worksheet.");
        }
    }
}