using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureCheckDemo
    {
        public static void Run()
        {
            const string inputFile = "SignedVbaProject.xlsm";
            const string copyFile = "SignedVbaProject_Copy.xlsm";

            // Ensure the input file exists; if not, create a simple macro‑enabled workbook
            if (!File.Exists(inputFile))
            {
                var wb = new Workbook();
                wb.Save(inputFile, SaveFormat.Xlsm);
            }

            // Load the workbook that (may) contain a VBA project
            var workbook = new Workbook(inputFile);

            // Determine whether the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + workbook.VbaProject.IsSigned);

            // If it is signed, verify whether the signature is valid
            if (workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA Project Signature Valid: " + workbook.VbaProject.IsValidSigned);
            }

            // Save the workbook to a macro‑enabled format – signature information is preserved
            workbook.Save(copyFile, SaveFormat.Xlsm);

            // Reload the workbook from the saved file to confirm the signature status remains unchanged
            var reloadedWorkbook = new Workbook(copyFile);

            Console.WriteLine("After reload - VBA Project Signed: " + reloadedWorkbook.VbaProject.IsSigned);
            Console.WriteLine("After reload - VBA Project Signature Valid: " + reloadedWorkbook.VbaProject.IsValidSigned);
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}