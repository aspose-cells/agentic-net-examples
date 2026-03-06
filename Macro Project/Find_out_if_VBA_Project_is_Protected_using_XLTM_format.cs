using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLTM file (macro-enabled template)
            string filePath = "sample.xltm";

            // Load the workbook (XLTM format is automatically detected)
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Output the protection status
            Console.WriteLine($"Is VBA Project Protected: {isProtected}");

            // Optional: also display whether it is locked for viewing
            Console.WriteLine($"Is VBA Project Locked for Viewing: {vbaProject.IslockedForViewing}");
        }
    }
}