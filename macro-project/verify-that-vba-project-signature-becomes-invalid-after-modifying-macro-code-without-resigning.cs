using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureInvalidAfterModificationDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Path to the signed workbook (replace with actual path)
            string signedWorkbookPath = "signed.xlsm";

            // Prevent FileNotFoundException
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"File not found: {signedWorkbookPath}");
                return;
            }

            try
            {
                // Load the workbook that already contains a signed VBA project
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Show initial signature status
                Console.WriteLine("Initial state:");
                Console.WriteLine("Is VBA Project Signed: " + workbook.VbaProject.IsSigned);
                Console.WriteLine("Is Signature Valid: " + workbook.VbaProject.IsValidSigned);

                // Ensure at least one module exists; add one if necessary
                if (workbook.VbaProject.Modules.Count == 0)
                {
                    // Use cast to VbaModuleType to avoid enum version issues
                    int newModuleIdx = workbook.VbaProject.Modules.Add((VbaModuleType)0, "Mod1");
                    workbook.VbaProject.Modules[newModuleIdx].Codes = "Sub Dummy()\nEnd Sub";
                }

                // Append a comment line to the first module's code
                VbaModule firstModule = workbook.VbaProject.Modules[0];
                firstModule.Codes += "\n' Modification made without re‑signing";

                // Save the modified workbook to a memory stream (preserving macro format)
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsm);
                    ms.Position = 0; // reset stream position for reading

                    // Reload the workbook from the stream
                    Workbook modifiedWorkbook = new Workbook(ms);

                    // Show signature status after modification
                    Console.WriteLine("\nAfter modification (without re‑signing):");
                    Console.WriteLine("Is VBA Project Signed: " + modifiedWorkbook.VbaProject.IsSigned);
                    Console.WriteLine("Is Signature Valid: " + modifiedWorkbook.VbaProject.IsValidSigned);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Runtime error: " + ex.Message);
            }
        }
    }
}