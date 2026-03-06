using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaSignatureFods
{
    static void Main()
    {
        // Path to the FODS (OpenDocument Spreadsheet) file
        string fodsPath = "sample.fods";

        // Load the workbook from the FODS file
        Workbook workbook = new Workbook(fodsPath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Output whether the VBA project is signed
        Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

        // If signed, output whether the signature is valid
        if (vbaProject.IsSigned)
        {
            Console.WriteLine("VBA Signature Valid: " + vbaProject.IsValidSigned);
        }
        else
        {
            Console.WriteLine("VBA project is not signed; no signature to validate.");
        }
    }
}