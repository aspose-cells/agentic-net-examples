using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaProjectSignatureValidationAfterCopy
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            const string sourcePath = "SignedSource.xlsm";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook that already contains a signed VBA project
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Verify that the source VBA project is signed and the signature is valid
            Console.WriteLine("Source - IsSigned: " + sourceWorkbook.VbaProject.IsSigned);
            Console.WriteLine("Source - IsValidSigned: " + sourceWorkbook.VbaProject.IsValidSigned);

            // Create a new empty macro-enabled workbook
            Workbook destWorkbook = new Workbook();
            using (MemoryStream tempStream = new MemoryStream())
            {
                destWorkbook.Save(tempStream, SaveFormat.Xlsm);
                tempStream.Position = 0;
                destWorkbook = new Workbook(tempStream);
            }

            // Copy the VBA project from the source workbook to the destination workbook
            destWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

            // Save the destination workbook to a memory stream and reload to simulate a fresh load
            using (MemoryStream copiedStream = new MemoryStream())
            {
                destWorkbook.Save(copiedStream, SaveFormat.Xlsm);
                copiedStream.Position = 0;

                Workbook reloadedWorkbook = new Workbook(copiedStream);

                // Validate the VBA project's signature in the reloaded workbook
                Console.WriteLine("After Copy - IsSigned: " + reloadedWorkbook.VbaProject.IsSigned);
                Console.WriteLine("After Copy - IsValidSigned: " + reloadedWorkbook.VbaProject.IsValidSigned);
            }
        }
    }
}