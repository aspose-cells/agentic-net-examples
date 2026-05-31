using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureRemoval
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a signed VBA project
            string signedPath = "SignedVbaProject.xlsm";

            // Load the signed workbook
            Workbook signedWorkbook = new Workbook(signedPath);

            // Verify that the VBA project is currently signed
            Console.WriteLine("Before removal - VBA project IsSigned: " + signedWorkbook.VbaProject.IsSigned);

            // Remove the digital signature from the workbook (this also clears the VBA project signature)
            signedWorkbook.RemoveDigitalSignature();

            // Save the workbook without the signature
            string unsignedPath = "UnsignedVbaProject.xlsm";
            signedWorkbook.Save(unsignedPath, SaveFormat.Xlsm);

            // Reload the saved workbook to verify the signature state
            Workbook unsignedWorkbook = new Workbook(unsignedPath);
            Console.WriteLine("After removal - VBA project IsSigned: " + unsignedWorkbook.VbaProject.IsSigned);
        }
    }
}