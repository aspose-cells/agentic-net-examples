using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file that contains a VBA project (e.g., .xlsm)
            string workbookPath = "SignedVbaWorkbook.xlsm";

            // Load the workbook from the file
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

            // If the VBA project is signed, verify the validity of the signature
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA Project Signature Valid: " + vbaProject.IsValidSigned);
            }

            // Optionally, check if the entire workbook is digitally signed
            Console.WriteLine("Workbook Digitally Signed: " + workbook.IsDigitallySigned);

            if (workbook.IsDigitallySigned)
            {
                // Retrieve the collection of digital signatures attached to the workbook
                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                // Iterate through each signature and display its validation status
                foreach (DigitalSignature signature in signatures)
                {
                    Console.WriteLine("Workbook Signature Valid: " + signature.IsValid);
                }
            }

            // No modifications are made, so saving is not required in this example.
        }
    }
}