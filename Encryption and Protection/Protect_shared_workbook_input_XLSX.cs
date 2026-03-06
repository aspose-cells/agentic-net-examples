using System;
using Aspose.Cells;

namespace ProtectSharedWorkbookExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that needs to be protected
            string inputPath = "input.xlsx";

            // Desired output path for the protected workbook
            string outputPath = "protected_shared.xlsx";

            // Password to protect the shared workbook
            string password = "myPassword";

            // Load the existing workbook
            Workbook wb = new Workbook(inputPath);

            // Protect the workbook as a shared workbook with the specified password
            wb.ProtectSharedWorkbook(password);

            // Save the protected workbook to the output file
            wb.Save(outputPath);
        }
    }
}