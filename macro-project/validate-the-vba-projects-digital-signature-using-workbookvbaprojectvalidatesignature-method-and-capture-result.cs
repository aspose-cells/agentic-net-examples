using System;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main()
        {
            // Load a workbook that contains a VBA project (macro-enabled file)
            string workbookPath = "SignedWorkbook.xlsm";
            Workbook workbook = new Workbook(workbookPath);

            // Determine whether the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            // Validate the digital signature of the VBA project
            // The IsValidSigned property reflects the result of the validation
            bool isSignatureValid = workbook.VbaProject.IsValidSigned;
            Console.WriteLine("VBA Project Signature Valid: " + isSignatureValid);
        }
    }
}