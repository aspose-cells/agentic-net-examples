using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    public class ValidateVbaSignatureInWorkbook
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string filePath = "sample.xlsm";

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook without specifying a format; Aspose.Cells will auto-detect.
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(filePath, loadOptions);

            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("VBA signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or not present.");
            }

            Console.WriteLine("Workbook digitally signed: " + workbook.IsDigitallySigned);
            if (workbook.IsDigitallySigned)
            {
                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();
                foreach (DigitalSignature sig in signatures)
                {
                    Console.WriteLine("Workbook signature valid: " + sig.IsValid);
                }
            }
        }
    }
}