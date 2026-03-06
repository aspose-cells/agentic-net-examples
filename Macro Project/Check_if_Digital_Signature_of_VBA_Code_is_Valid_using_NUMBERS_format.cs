using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    public class VbaDigitalSignatureCheckDemo
    {
        public static void Run()
        {
            // Path to the Excel file that contains a VBA project
            string workbookPath = "SignedVbaWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed (1 = true, 0 = false)
            int isSigned = vbaProject.IsSigned ? 1 : 0;
            Console.WriteLine("VBA Project IsSigned (numeric): " + isSigned);

            // If signed, check whether the signature is valid (1 = true, 0 = false)
            int isValidSigned = vbaProject.IsValidSigned ? 1 : 0;
            Console.WriteLine("VBA Project IsValidSigned (numeric): " + isValidSigned);

            // Verify digital signatures attached to the workbook itself
            DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

            if (signatures != null)
            {
                int index = 1;
                foreach (DigitalSignature sig in signatures)
                {
                    int sigIsValid = sig.IsValid ? 1 : 0;
                    Console.WriteLine($"Digital Signature #{index} IsValid (numeric): {sigIsValid}");
                    index++;
                }

                if (index == 1)
                {
                    Console.WriteLine("No digital signatures found in the workbook.");
                }
            }
            else
            {
                Console.WriteLine("No digital signatures collection available.");
            }

            // Save the workbook to a new file to demonstrate that the signature information persists
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0; // Reset stream position for reading

                // Reload and re‑check to confirm persistence
                Workbook reloaded = new Workbook(ms);
                VbaProject reloadedVba = reloaded.VbaProject;
                Console.WriteLine("After reload - IsSigned (numeric): " + (reloadedVba.IsSigned ? 1 : 0));
                Console.WriteLine("After reload - IsValidSigned (numeric): " + (reloadedVba.IsValidSigned ? 1 : 0));
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaDigitalSignatureCheckDemo.Run();
        }
    }
}