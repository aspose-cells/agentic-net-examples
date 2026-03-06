using System;
using Aspose.Cells;

namespace AsposeCellsUnprotectDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (provide via command line or use defaults)
            string inputPath = args.Length > 0 ? args[0] : "protected_input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "unprotected_output.xlsx";

            // Optional password for the worksheet; if not supplied the worksheet is assumed to be protected without a password
            string password = args.Length > 2 ? args[2] : null;

            // Load the workbook from the input XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the appropriate overload
            if (string.IsNullOrEmpty(password))
            {
                // Worksheet protected without a password
                worksheet.Unprotect();
            }
            else
            {
                // Worksheet protected with a password
                worksheet.Unprotect(password);
            }

            // Save the unprotected workbook to the output file
            workbook.Save(outputPath);
        }
    }
}