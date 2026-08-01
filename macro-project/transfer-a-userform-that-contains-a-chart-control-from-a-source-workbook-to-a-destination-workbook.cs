// Title: Copy a VBA UserForm with Chart Control between .xlsm workbooks using Aspose.Cells for .NET (C#)
// Description: Shows how to load a macro‑enabled workbook, duplicate its entire VBA project—including modules, UserForms and embedded chart controls—into a new workbook and save it as an .xlsm file with Aspose.Cells VbaProject.Copy in C#.
// Keywords: Aspose.Cells | C# | VbaProject.Copy | copy VBA UserForm | transfer chart control | macro-enabled workbook | xlsm automation | VBA project cloning | Excel UI migration | UserForm duplication
// Common Searches: asp... copy userform asp... .net c# | how to copy VBA userform with chart using Aspose.Cells | transfer VBA project between .xlsm files c# | duplicate macro-enabled workbook Aspose.Cells | copy chart control on userform programmatically
// Developer Intent: Duplicate the VBA UserForm (including its chart control) from a source .xlsm file to a new macro‑enabled workbook using Aspose.Cells.
// Use Cases: Reuse custom Excel UI across multiple reports without rebuilding forms. | Automate creation of workbook templates that contain pre‑built user forms and charts. | Migrate legacy add‑ins to new files while preserving all VBA components. | Generate batch copies of macro‑enabled workbooks for distribution to end users.
// AI Prompts: Write C# code that copies only a selected UserForm from one .xlsm workbook to another using Aspose.Cells. | Show how to verify that a chart control inside a transferred UserForm loads data correctly after copying. | Create robust error handling for cases where the source workbook lacks a VBA project or the specified UserForm. | Explain how to copy VBA modules separately while excluding certain forms with Aspose.Cells. | Provide steps to preserve digital signatures after cloning a VBA project.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace TransferUserFormExample
{
    // Shows how to load a macro‑enabled workbook, duplicate its entire VBA project—including modules, UserForms and embedded chart controls—into a new workbook and save it as an .xlsm file with Aspose.Cells VbaProject.Copy in C#.
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string sourcePath = "SourceWithUserForm.xlsm";
                const string destPath = "DestinationWithUserForm.xlsm";

                // Ensure the source workbook exists before loading.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                    return;
                }

                // Load the source workbook (must be macro‑enabled to retain VBA).
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook.
                Workbook destinationWorkbook = new Workbook();

                // Copy the entire VBA project (modules, user forms, etc.) to the destination.
                destinationWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

                // Save the destination workbook as a macro‑enabled file.
                destinationWorkbook.Save(destPath, SaveFormat.Xlsm);

                Console.WriteLine($"UserForm transferred successfully to {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
