// Title: Check VBA project digital signature after copying an .xlsm workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a signed .xlsm file, saves it to a temporary location, reloads it, and uses Workbook.VbaProject.IsSigned to confirm the signature remains. | Create a .NET method that copies a macro-enabled workbook and returns a boolean indicating whether the VBA project's digital signature is still present. | Generate a C# snippet that validates the integrity of a VBA project's signature after saving and reloading the workbook using Aspose.Cells.
// Common Searches: aspnet verify VBA project signature after workbook copy | how to check if VBA macro is still signed using Aspose.Cells C# | C# detect lost digital signature in .xlsm after save with Aspose.Cells | validate VBA project IsSigned flag after exporting macro-enabled Excel file | preserve VBA digital signature when copying .xlsm using Aspose.Cells .NET
// Tags: VbaProject.IsSigned verification | copy signed .xlsm workbook Aspose.Cells | validate VBA digital signature .NET | temporary file handling Aspose.Cells | macro workbook integrity check C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads a signed macro-enabled .xlsm workbook, saves it to a temporary file, reloads it, and uses the VbaProject.IsSigned property to verify that the VBA project's digital signature remains intact.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsm";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the original workbook that contains a signed VBA project
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Save the workbook (including its VBA project) to a temporary file
            string tempDestPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
            sourceWorkbook.Save(tempDestPath, SaveFormat.Xlsm);

            // Load the temporary workbook as the destination workbook
            Workbook destinationWorkbook;
            try
            {
                destinationWorkbook = new Workbook(tempDestPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load temporary workbook: {ex.Message}");
                return;
            }

            // Clean up the temporary file
            try
            {
                if (File.Exists(tempDestPath))
                    File.Delete(tempDestPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not delete temporary file: {ex.Message}");
            }

            // Check whether the VBA project is still signed after copy
            bool signatureIsValid = false;
            try
            {
                var vbaProject = destinationWorkbook.VbaProject;
                if (vbaProject != null)
                {
                    // Aspose.Cells provides only IsSigned flag; assume true means signature is present
                    signatureIsValid = vbaProject.IsSigned;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while inspecting VBA project: {ex.Message}");
            }

            Console.WriteLine("VBA project signature present after copy: " + signatureIsValid);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
