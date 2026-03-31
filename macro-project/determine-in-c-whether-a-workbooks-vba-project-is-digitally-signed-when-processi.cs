using System;
using Aspose.Cells;

class CheckVbaSignatureInFods
{
    static void Main()
    {
        // Load the FODS workbook
        Workbook workbook = new Workbook("input.fods");

        // Verify that the workbook contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Determine if the VBA project is digitally signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project is signed: " + isSigned);

            // If signed, optionally check whether the signature is valid
            if (isSigned)
            {
                Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}