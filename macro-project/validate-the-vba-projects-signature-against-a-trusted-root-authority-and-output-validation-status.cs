using System;
using Aspose.Cells;

class ValidateVbaSignature
{
    static void Main()
    {
        // Load the workbook that contains a VBA project.
        Workbook workbook = new Workbook("signedWorkbook.xlsm");

        // Check whether the VBA project is signed.
        if (workbook.VbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");

            // Validate the signature against the trusted root authority.
            bool isSignatureValid = workbook.VbaProject.IsValidSigned;
            Console.WriteLine("Signature valid: " + isSignatureValid);
        }
        else
        {
            Console.WriteLine("VBA project is not signed.");
        }
    }
}