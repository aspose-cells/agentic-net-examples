using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    class XlsxToCsv
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Path for the resulting CSV file
            string destPath = "output.csv";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Configure CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions
            {
                // Preserve delimiter placeholders for completely blank rows
                KeepSeparatorsForBlankRow = true,

                // Do not trim leading blank rows/columns so row alignment stays intact
                TrimLeadingBlankRowAndColumn = false,

                // Use UTF‑8 encoding (optional, but common for CSV)
                Encoding = Encoding.UTF8
            };

            // Save the workbook as CSV using the configured options
            workbook.Save(destPath, csvOptions);

            Console.WriteLine($"Workbook converted to CSV successfully: {destPath}");
        }
    }
}