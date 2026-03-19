using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSB workbook to be examined
            string workbookPath = "sample.xlsb";

            // Load the workbook (XLSB format)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Output the result
            Console.WriteLine($"VBA Project Protected: {isProtected}");
        }
    }
}