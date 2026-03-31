using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace VbaProjectProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro-enabled workbook (XLSM) to be examined.
            // Provide the file path as the first command‑line argument or replace with a literal.
            string workbookPath = args.Length > 0 ? args[0] : "sample.xlsm";

            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is protected.
            bool isProtected = vbaProject.IsProtected;

            // Determine if the VBA project is locked for viewing.
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Output the results.
            Console.WriteLine($"VBA Project Protected: {isProtected}");
            Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");

            // Optional: If the project is protected, you can attempt to validate a password.
            // Replace "yourPassword" with the password you want to test.
            if (isProtected)
            {
                string testPassword = "yourPassword";
                bool passwordValid = vbaProject.ValidatePassword(testPassword);
                Console.WriteLine($"Password validation result for \"{testPassword}\": {passwordValid}");
            }
        }
    }
}