using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ValidateVbaSignature
{
    static void Main()
    {
        // Load a workbook that contains a VBA project.
        Workbook workbook = new Workbook("SignedWorkbook.xlsm");

        // Access the VBA project.
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is signed.
        if (vbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");

            // Validate the signature against the trusted root authority.
            // The IsValidSigned property returns true if the signature is valid.
            Console.WriteLine("Signature valid: " + vbaProject.IsValidSigned);
        }
        else
        {
            Console.WriteLine("VBA project is not signed.");
        }
    }
}