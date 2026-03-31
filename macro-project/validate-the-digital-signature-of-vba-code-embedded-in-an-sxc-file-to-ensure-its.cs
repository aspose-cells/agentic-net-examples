using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            VbaSignatureValidator.Run();
        }
    }

    public class VbaSignatureValidator
    {
        public static void Run()
        {
            // Load the SXC workbook that may contain a VBA project
            Workbook workbook = new Workbook("sample.sxc");

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Ensure the VBA project exists
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Check whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            // If signed, verify the validity of the signature
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isValid);
            }
            else
            {
                Console.WriteLine("VBA Project is not signed; no signature to validate.");
            }
        }
    }
}