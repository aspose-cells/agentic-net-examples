using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook("sample.xlsm");

            // Ensure the workbook has a VBA project
            if (workbook.VbaProject != null)
            {
                // Determine whether the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA Project Signed: " + isSigned);

                if (isSigned)
                {
                    // If signed, check whether the signature is valid
                    bool isValid = workbook.VbaProject.IsValidSigned;
                    Console.WriteLine("VBA Signature Valid: " + isValid);
                }
                else
                {
                    Console.WriteLine("VBA project is not signed; no signature to validate.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
    }
}