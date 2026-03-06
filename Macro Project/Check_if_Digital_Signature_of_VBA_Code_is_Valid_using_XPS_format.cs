using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaSignatureXpsDemo
{
    static void Main()
    {
        // Load an Excel workbook that contains a VBA project
        Workbook workbook = new Workbook("SignedVbaWorkbook.xlsm");

        // Determine whether the VBA project is signed
        bool isSigned = workbook.VbaProject.IsSigned;
        Console.WriteLine("VBA Project Signed: " + isSigned);

        // If it is signed, verify the validity of the signature
        if (isSigned)
        {
            bool isValid = workbook.VbaProject.IsValidSigned;
            Console.WriteLine("VBA Signature Valid: " + isValid);
        }

        // Save the workbook as XPS (visual representation of the sheet)
        workbook.Save("SignedVbaWorkbook.xps", SaveFormat.Xps);
        Console.WriteLine("Workbook saved as XPS.");
    }
}