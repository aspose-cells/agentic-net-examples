using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel workbook that contains a VBA project
            string workbookPath = "SignedVbaWorkbook.xlsm";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(workbookPath);

            // Retrieve the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project exists and is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Check whether the digital signature of the VBA project is valid
                bool isSignatureValid = vbaProject.IsValidSigned;
                Console.WriteLine("Signature valid: " + isSignatureValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }
}