using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the path of the Excel file to inspect
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsVbaSignatureCheck <excel-file-path>");
                return;
            }

            string excelPath = args[0];

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(excelPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed and if the signature is valid
            bool isSigned = vbaProject.IsSigned;               // uses VbaProject.IsSigned property
            bool isValidSigned = vbaProject.IsValidSigned;     // uses VbaProject.IsValidSigned property

            // Prepare CSV line: FileName,IsSigned,IsValidSigned
            string csvLine = $"{Path.GetFileName(excelPath)},{isSigned},{isValidSigned}";

            // Output CSV to console (could be redirected to a file if needed)
            Console.WriteLine(csvLine);
        }
    }
}