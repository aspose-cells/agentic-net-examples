using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (must be a macro-enabled file, e.g., .xlsm)
            string sourcePath = "SampleWithVba.xlsm";

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project exists and is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }

            // Save the workbook in SXC format (OpenOffice Spreadsheet)
            // Note: SXC does not preserve VBA macros, but the check is performed before saving.
            string outputPath = "ConvertedWorkbook.sxc";
            workbook.Save(outputPath, SaveFormat.Sxc);

            Console.WriteLine("Workbook saved as SXC to: " + Path.GetFullPath(outputPath));
        }
    }
}