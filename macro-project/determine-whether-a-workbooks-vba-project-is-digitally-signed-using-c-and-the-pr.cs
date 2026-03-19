using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class CheckVbaProjectSignatureInPrn
    {
        public static void Run()
        {
            // Path to the PRN file that contains the workbook
            string prnPath = "sample.prn";

            if (!File.Exists(prnPath))
            {
                Console.WriteLine($"File not found: {prnPath}");
                return;
            }

            // Load the workbook from the PRN file with default load options (auto-detect format)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(prnPath, loadOptions);

            // Get the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is digitally signed
            bool isSigned = vbaProject != null && vbaProject.IsSigned;

            Console.WriteLine("VBA project is signed: " + isSigned);

            // If signed, optionally display whether the signature is valid
            if (isSigned && vbaProject != null)
            {
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CheckVbaProjectSignatureInPrn.Run();
        }
    }
}