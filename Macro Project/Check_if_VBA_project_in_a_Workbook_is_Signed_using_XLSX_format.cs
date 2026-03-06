using System;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Load the workbook (XLSX format)
        Workbook workbook = new Workbook("sample.xlsx");

        // Verify that the workbook contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Check if the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project is signed: " + isSigned);

            // If signed, also display whether the signature is valid
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