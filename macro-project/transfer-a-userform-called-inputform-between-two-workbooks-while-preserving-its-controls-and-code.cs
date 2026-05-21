using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class TransferUserForm
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWithForm.xlsm";
            const string destPath = "DestinationWithForm.xlsm";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook that contains the UserForm named "InputForm"
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Copy the entire VBA project (modules, userforms, references, etc.) from source to destination
            destinationWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

            // Ensure that any external VBA references are also copied (optional, but safe)
            destinationWorkbook.VbaProject.References.Copy(sourceWorkbook.VbaProject.References);

            // Save the destination workbook as a macro‑enabled file to preserve the UserForm
            destinationWorkbook.Save(destPath, SaveFormat.Xlsm);

            Console.WriteLine($"VBA project successfully transferred to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}