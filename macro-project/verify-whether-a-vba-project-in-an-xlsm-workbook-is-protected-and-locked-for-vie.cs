using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro-enabled workbook (XLSM) you want to inspect
            string workbookPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Check if the VBA project is locked for viewing
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Output the results
            Console.WriteLine($"VBA Project Protected: {isProtected}");
            Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
        }
    }
}