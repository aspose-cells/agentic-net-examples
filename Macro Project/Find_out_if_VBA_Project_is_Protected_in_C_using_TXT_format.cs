using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro-enabled workbook (XLSM) to be examined
            string workbookPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Prepare the output text
            string output = $"VBA Project Protected: {isProtected}";

            // Write the result to a TXT file
            string txtPath = "VbaProtectionStatus.txt";
            File.WriteAllText(txtPath, output);

            // Optional: display result in console
            Console.WriteLine(output);
            Console.WriteLine($"Result written to: {txtPath}");
        }
    }
}