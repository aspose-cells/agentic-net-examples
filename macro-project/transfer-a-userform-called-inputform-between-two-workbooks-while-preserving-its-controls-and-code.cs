// Title: Copy a VBA UserForm (InputForm) between macro‑enabled workbooks using Aspose.Cells for .NET
// Description: Loads a source .xlsm file, creates a new workbook, copies the entire VBA project—including the InputForm userform, its controls, and code—via VbaProject.Copy, and saves the destination as a macro‑enabled file. Includes basic file‑existence checking and error reporting.
// Keywords: Aspose.Cells | C# copy VBA project | transfer VBA UserForm | InputForm | macro-enabled workbook | VbaProject.Copy | .NET Excel automation | preserve VBA controls | Excel VBA form migration
// Common Searches: copy VBA UserForm between .xlsm files Aspose.Cells | Aspose.Cells transfer UserForm InputForm .NET | how to duplicate VBA project in C# Excel | preserve VBA controls when moving workbooks | macro‑enabled workbook copy using Aspose.Cells
// Developer Intent: Duplicate the VBA project that contains the InputForm userform from a source workbook to a new macro‑enabled workbook using Aspose.Cells for .NET.
// Use Cases: Deploy a standard InputForm to multiple generated reports while retaining all macro functionality. | Create a clean workbook that inherits macros and userforms from a template for consistent data entry. | Automate the distribution of a custom VBA userform across a batch of workbooks in a deployment pipeline.
// AI Prompts: Generate C# code with Aspose.Cells that copies only the InputForm userform, leaving other VBA modules untouched. | Add robust error handling to verify the presence of InputForm before copying and log detailed status messages. | Explain how to adjust reference paths inside the transferred InputForm code after moving it to a new workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUserFormTransfer
{
    // Loads a source .xlsm file, creates a new workbook, copies the entire VBA project—including the InputForm userform, its controls, and code—via VbaProject.Copy, and saves the destination as a macro‑enabled file. Includes basic file‑existence checking and error reporting.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceWithForm.xlsm";
                const string destinationPath = "DestinationWithInputForm.xlsm";

                // Verify that the source workbook exists before attempting to load it
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                    return;
                }

                // Load the source workbook that contains the UserForm "InputForm"
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new empty workbook for the destination
                Workbook destinationWorkbook = new Workbook();

                // Copy the entire VBA project (including user forms, modules, etc.) from source to destination
                destinationWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

                // Save the destination workbook as a macro‑enabled file
                destinationWorkbook.Save(destinationPath, SaveFormat.Xlsm);

                Console.WriteLine("UserForm 'InputForm' transferred successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
