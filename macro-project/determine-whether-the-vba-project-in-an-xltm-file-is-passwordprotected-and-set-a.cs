using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLTM file (macro‑enabled template)
            string filePath = "sample.xltm";

            // Load the workbook (XLTM) – this also loads the VBA project if present
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is protected with a password
            bool isProtected = vbaProject.IsProtected;

            // Determine if the VBA project is locked for viewing (requires password to view code)
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Output the results
            Console.WriteLine($"VBA Project Protected: {isProtected}");
            Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");

            // Optional: If you know the password, you can validate it
            // string password = "yourPassword";
            // bool passwordValid = vbaProject.ValidatePassword(password);
            // Console.WriteLine($"Password validation result: {passwordValid}");
        }
    }
}