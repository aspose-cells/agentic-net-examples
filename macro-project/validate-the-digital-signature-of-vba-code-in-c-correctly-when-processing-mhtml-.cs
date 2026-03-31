using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            VbaSignatureValidator.Run();
        }
    }

    public static class VbaSignatureValidator
    {
        public static void Run()
        {
            // Path to the MHTML file that contains an Excel workbook with VBA project
            string mhtmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.mht");

            if (!File.Exists(mhtmlPath))
            {
                Console.WriteLine($"File not found: {mhtmlPath}");
                // Create an empty workbook to demonstrate loading/saving flow
                Workbook emptyWb = new Workbook();
                string emptyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "empty.xlsm");
                emptyWb.Save(emptyPath, SaveFormat.Xlsm);
                Console.WriteLine($"Created empty workbook at: {emptyPath}");
                return;
            }

            // Load the workbook from the MHTML file (format detection is automatic)
            Workbook workbook = new Workbook(mhtmlPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Validate the digital signature of the VBA project
                bool isSignatureValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA signature is valid: " + isSignatureValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }

            // (Optional) Save the workbook to a memory stream to verify that the signature state persists
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0; // Reset stream position for reading

                // Reload the workbook from the memory stream
                Workbook reloadedWorkbook = new Workbook(ms);
                VbaProject reloadedVba = reloadedWorkbook.VbaProject;

                Console.WriteLine("After reload - VBA signed: " + (reloadedVba != null && reloadedVba.IsSigned));
                Console.WriteLine("After reload - VBA signature valid: " + (reloadedVba != null && reloadedVba.IsValidSigned));
            }
        }
    }
}