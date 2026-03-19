using System;
using Aspose.Cells;

class ValidateVbaSignature
{
    static void Main()
    {
        // Load the ODS workbook
        Workbook workbook = new Workbook("sample.ods");

        // Determine if the workbook contains any VBA macros
        if (workbook.HasMacro)
        {
            Console.WriteLine("Workbook contains VBA macros.");

            // Check whether the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            if (isSigned)
            {
                // Verify that the signature is valid (i.e., not tampered)
                bool isValid = workbook.VbaProject.IsValidSigned;
                Console.WriteLine("VBA signature valid: " + isValid);
            }
        }
        else
        {
            Console.WriteLine("Workbook does not contain VBA macros.");
        }

        // Optionally, report the overall digital signature status of the workbook
        Console.WriteLine("Workbook digitally signed: " + workbook.IsDigitallySigned);
    }
}