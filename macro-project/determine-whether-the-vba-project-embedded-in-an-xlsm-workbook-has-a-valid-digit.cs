using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignature
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the macro-enabled workbook
            string filePath = "sample.xlsm";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Ensure the workbook contains a VBA project (macro)
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Determine whether the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA Project Signed: " + isSigned);

                // If signed, check if the signature is valid
                if (isSigned)
                {
                    bool isValid = workbook.VbaProject.IsValidSigned;
                    Console.WriteLine("VBA Project Signature Valid: " + isValid);
                }
                else
                {
                    Console.WriteLine("VBA Project is not signed; no signature to validate.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
    }
}