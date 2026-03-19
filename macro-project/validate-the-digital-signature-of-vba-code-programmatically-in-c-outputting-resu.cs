using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled) to be validated.
            // You can pass the path as a command‑line argument; otherwise a default is used.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsm";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File not found:\t" + filePath);
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed.
            bool isSigned = vbaProject.IsSigned;

            // Determine whether the signature (if present) is valid.
            bool isValidSigned = vbaProject.IsValidSigned;

            // Optional: check if the whole workbook is digitally signed.
            bool workbookSigned = workbook.IsDigitallySigned;

            // Output the results in TAB‑delimited format.
            // Columns: FilePath   IsVbaSigned   IsVbaSignatureValid   IsWorkbookDigitallySigned
            Console.WriteLine(
                $"{filePath}\t{isSigned}\t{isValidSigned}\t{workbookSigned}"
            );
        }
    }
}