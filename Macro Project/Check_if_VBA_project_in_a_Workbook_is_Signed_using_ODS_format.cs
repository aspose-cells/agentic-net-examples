using System;
using Aspose.Cells;

class CheckVbaSignedOds
{
    static void Main()
    {
        // Path to the ODS file
        string odsPath = "sample.ods";

        // Load the workbook from ODS format
        Workbook workbook = new Workbook(odsPath);

        // Verify that the workbook contains a VBA project (macro)
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Check if the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // If signed, optionally display whether the signature is valid
            if (isSigned)
            {
                Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}