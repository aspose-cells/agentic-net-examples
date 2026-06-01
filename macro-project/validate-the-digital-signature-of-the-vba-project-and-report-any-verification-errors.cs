using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ValidateVbaSignature
{
    static void Main()
    {
        // Load the workbook that contains a VBA project
        Workbook workbook = new Workbook("sample.xlsm");

        // Determine whether the VBA project is signed
        if (workbook.VbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");

            // Verify if the signature is valid
            bool isValid = workbook.VbaProject.IsValidSigned;
            Console.WriteLine("Signature valid: " + isValid);

            // Report any verification errors
            if (!isValid)
            {
                Console.WriteLine("Error: VBA project signature verification failed.");
            }
        }
        else
        {
            Console.WriteLine("VBA project is not signed.");
        }
    }
}