using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaUnlockDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a VBA project locked for viewing
            string inputPath = "protected.xlsm";

            // Path where the unlocked workbook will be saved
            string outputPath = "unlocked.xlsm";

            // Password that should unlock the VBA project
            string password = "yourPasswordHere";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                VbaProject vbaProject = workbook.VbaProject;

                // Report current protection status
                Console.WriteLine($"Is VBA Project Protected: {vbaProject.IsProtected}");

                // Validate the supplied password
                bool isPasswordValid = vbaProject.ValidatePassword(password);
                Console.WriteLine($"Password validation result: {isPasswordValid}");

                if (isPasswordValid)
                {
                    // Unprotect (unlock) the VBA project
                    // Passing false for isLockedForViewing and null for password removes protection
                    vbaProject.Protect(false, null);
                    Console.WriteLine("VBA project unlocked successfully.");
                }
                else
                {
                    Console.WriteLine("Failed to unlock VBA project: invalid password.");
                }

                // Save the workbook after attempting to unlock
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}