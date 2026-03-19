using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel workbook that may contain a VBA project
            string workbookPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // Determine if the VBA project is signed
                bool isSigned = vbaProject.IsSigned;
                Console.WriteLine("VBA Project Signed: " + isSigned);

                // If signed, check whether the signature is valid
                if (isSigned)
                {
                    bool isValid = vbaProject.IsValidSigned;
                    Console.WriteLine("VBA Project Signature Valid: " + isValid);
                }
                else
                {
                    Console.WriteLine("VBA Project is not signed; signature validity check is not applicable.");
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }

            // (Optional) Save the workbook to preserve any changes
            // workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
    }
}