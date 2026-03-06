using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationDemo
    {
        public static void Run()
        {
            // Path to the macro-enabled workbook that contains a signed VBA project
            string inputPath = "SignedVbaWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

            // If signed, check whether the signature is valid
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA Project Signature Valid: " + vbaProject.IsValidSigned);
            }

            // Save to a memory stream to ensure the signature persists after saving
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0; // Reset stream position for reading

                // Reload the workbook from the memory stream
                Workbook reloadedWorkbook = new Workbook(ms);
                VbaProject reloadedVbaProject = reloadedWorkbook.VbaProject;

                // Verify signature status after reload
                Console.WriteLine("After Reload - VBA Project Signed: " + reloadedVbaProject.IsSigned);
                Console.WriteLine("After Reload - VBA Project Signature Valid: " + reloadedVbaProject.IsValidSigned);
            }
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}