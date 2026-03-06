using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignedInXltxDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string filePath = "template.xltx";

            // Create a macro‑free XLTX template
            Workbook wbCreate = new Workbook();
            wbCreate.Save(filePath, SaveFormat.Xltx);

            // Load the XLTX template
            Workbook workbook = new Workbook(filePath);

            // Verify whether the workbook contains any VBA/macros
            Console.WriteLine("Workbook has macro: " + workbook.HasMacro);

            // Check if the (non‑existent) VBA project is signed
            bool isSigned = false;
            if (workbook.VbaProject != null)
            {
                isSigned = workbook.VbaProject.IsSigned;
            }
            Console.WriteLine("VBA project is signed: " + isSigned);

            // If it were signed, also display the validity of the signature
            if (isSigned && workbook.VbaProject != null)
            {
                Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
            }

            // Clean up the temporary file
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }
}