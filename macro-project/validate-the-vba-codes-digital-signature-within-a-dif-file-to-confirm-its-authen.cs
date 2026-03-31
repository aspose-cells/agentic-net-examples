using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    public class ValidateVbaSignature
    {
        public static void Run()
        {
            // Path to the DIF file that contains a VBA project
            string difPath = "sample.dif";

            // Load the workbook from the DIF file
            Workbook workbook = new Workbook(difPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

            // If signed, verify the validity of the signature
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA Project Signature Valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed; no signature to validate.");
            }

            // Optional: Save to a memory stream in XLSM format (required for VBA) and reload
            // to confirm that the signature information persists.
            using (MemoryStream ms = new MemoryStream())
            {
                // Save the workbook (including VBA) to the stream
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0; // Reset stream position for reading

                // Reload the workbook from the stream
                Workbook reloadedWorkbook = new Workbook(ms);
                VbaProject reloadedVba = reloadedWorkbook.VbaProject;

                // Output the signature status after reload
                Console.WriteLine("After reload - Signed: " + reloadedVba.IsSigned);
                Console.WriteLine("After reload - Signature Valid: " + reloadedVba.IsValidSigned);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateVbaSignature.Run();
        }
    }
}