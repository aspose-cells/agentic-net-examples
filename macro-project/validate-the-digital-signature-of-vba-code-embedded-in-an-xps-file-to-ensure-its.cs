using System;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains the VBA project (replace with actual file path)
            string workbookPath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            var vbaProject = workbook.VbaProject;

            // Output whether the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

            // If signed, validate the digital signature
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA Signature Valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("No VBA digital signature present.");
            }

            // Save a copy to preserve the signature (optional)
            workbook.Save("CopyWithSignature.xlsm");
        }
    }
}