using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureVerificationDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the Excel file that contains the VBA project.
            // The VBA project may have an associated .txt signature file.
            string workbookPath = "SignedVbaWorkbook.xlsm";

            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed.
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // If signed, also verify whether the signature is valid.
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA signature valid: " + isValid);
            }
        }
    }
}