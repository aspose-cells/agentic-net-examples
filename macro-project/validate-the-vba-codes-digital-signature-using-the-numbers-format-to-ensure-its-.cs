using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ValidateVbaSignatureDemo
    {
        public static void Run()
        {
            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook("SignedVbaWorkbook.xlsm");

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Check if the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            if (isSigned)
            {
                // Validate the digital signature of the VBA project
                bool isValidSignature = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isValidSignature);
            }
            else
            {
                Console.WriteLine("VBA Project is not signed; no signature to validate.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateVbaSignatureDemo.Run();
        }
    }
}