using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    public class Program
    {
        public static void Main()
        {
            // Path to the Excel file that contains a VBA project (xlsm)
            string inputPath = "SignedVbaWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            // If signed, verify whether the signature is valid
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isValid);
            }
            else
            {
                Console.WriteLine("VBA Project is not signed, therefore no signature to validate.");
            }

            // Demonstrate that the signature information persists after saving and reloading
            using (MemoryStream ms = new MemoryStream())
            {
                // Save the workbook to a memory stream (preserves the signature)
                workbook.Save(ms, SaveFormat.Xlsm);

                // Reset stream position for reading
                ms.Position = 0;

                // Reload the workbook from the memory stream
                Workbook reloadedWorkbook = new Workbook(ms);
                VbaProject reloadedVba = reloadedWorkbook.VbaProject;

                // Output the signature status of the reloaded workbook
                Console.WriteLine("After reload - VBA Project Signed: " + reloadedVba.IsSigned);
                Console.WriteLine("After reload - VBA Project Signature Valid: " + reloadedVba.IsValidSigned);
            }
        }
    }
}