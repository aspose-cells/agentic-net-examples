using System;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Path to the workbook that may contain a VBA project
        string workbookPath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Verify that the workbook actually contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Check if the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // If signed, check whether the signature is valid
            if (isSigned)
            {
                bool isValid = workbook.VbaProject.IsValidSigned;
                Console.WriteLine("Signature valid: " + isValid);
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}