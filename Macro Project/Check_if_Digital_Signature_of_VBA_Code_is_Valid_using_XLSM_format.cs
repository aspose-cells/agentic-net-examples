using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignature
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Load the XLSM workbook that contains a VBA project
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

            // Save to a memory stream and reload to confirm the signature persists
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0; // Reset stream position for reading

                Workbook reloaded = new Workbook(ms);
                Console.WriteLine("After reload - Signed: " + reloaded.VbaProject.IsSigned);
                Console.WriteLine("After reload - Valid: " + reloaded.VbaProject.IsValidSigned);
            }
        }
    }
}