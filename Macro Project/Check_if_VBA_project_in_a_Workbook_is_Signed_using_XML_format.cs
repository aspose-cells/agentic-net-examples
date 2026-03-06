using System;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Load a macro-enabled workbook (XLSM) that may contain a VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Determine whether the VBA project is signed
        bool isSigned = workbook.VbaProject.IsSigned;
        Console.WriteLine("VBA project is signed: " + isSigned);

        // If signed, optionally verify that the signature is valid
        if (isSigned)
        {
            bool isValid = workbook.VbaProject.IsValidSigned;
            Console.WriteLine("Signature is valid: " + isValid);
        }
    }
}