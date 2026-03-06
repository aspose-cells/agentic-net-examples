using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLS file (Excel 97-2003 format) that may contain a VBA project
            string filePath = "sample.xls";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Output the result
            Console.WriteLine($"Is VBA Project Protected ({filePath}): {isProtected}");

            // Optional: also display whether the project is locked for viewing
            Console.WriteLine($"Is VBA Project Locked for Viewing: {vbaProject.IslockedForViewing}");
        }
    }
}