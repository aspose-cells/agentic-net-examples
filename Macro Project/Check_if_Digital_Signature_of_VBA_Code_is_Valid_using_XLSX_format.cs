using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            VbaSignatureValidationDemo.Run();
        }
    }

    public static class VbaSignatureValidationDemo
    {
        public static void Run()
        {
            // Path to the macro‑enabled workbook (XLSM) that may contain a signed VBA project
            string workbookPath = Path.Combine(Environment.CurrentDirectory, "SignedVbaProject.xlsm");

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found at: {workbookPath}");
                return;
            }

            // Load the workbook from the file system
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
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
                Console.WriteLine("VBA Project is not signed; signature validation not applicable.");
            }

            // Save the workbook to verify that the signature state persists after saving
            string outputPath = Path.Combine(Environment.CurrentDirectory, "SignedVbaProject_Copy.xlsm");
            workbook.Save(outputPath, SaveFormat.Xlsm);

            // Reload the saved copy and re‑check the signature status
            Workbook reloadedWorkbook = new Workbook(outputPath);
            VbaProject reloadedVba = reloadedWorkbook.VbaProject;
            Console.WriteLine("After reload - VBA Project Signed: " + reloadedVba.IsSigned);
            Console.WriteLine("After reload - VBA Project Signature Valid: " + reloadedVba.IsValidSigned);
        }
    }
}