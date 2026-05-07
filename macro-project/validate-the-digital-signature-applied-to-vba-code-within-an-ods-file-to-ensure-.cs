using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ValidateVbaSignature
{
    static void Main()
    {
        // Load the ODS workbook that may contain a VBA project
        Workbook workbook = new Workbook("sample.ods");

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is present and signed
        if (vbaProject != null && vbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");

            // Check whether the signature is valid (authentic and untampered)
            bool isValid = vbaProject.IsValidSigned;
            Console.WriteLine("Signature valid: " + isValid);
        }
        else
        {
            Console.WriteLine("VBA project is not signed or does not exist.");
        }
    }
}