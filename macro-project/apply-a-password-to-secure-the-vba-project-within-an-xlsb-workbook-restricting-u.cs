using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSB supports macros)
            Workbook workbook = new Workbook();

            // Ensure the workbook has a VBA project by saving as a macro-enabled format
            // (XLSB is a macro-enabled binary format)
            // No need to add modules for protection demonstration
            // Protect the VBA project and lock it for viewing with a password
            workbook.VbaProject.Protect(true, "SecurePassword123");

            // Save the workbook with the protected VBA project
            workbook.Save("SecureVbaProject.xlsb", SaveFormat.Xlsb);

            // Optional: Verify protection status
            Console.WriteLine("VBA Project Protected: " + workbook.VbaProject.IsProtected);
            Console.WriteLine("VBA Project Locked for Viewing: " + workbook.VbaProject.IslockedForViewing);
        }
    }
}