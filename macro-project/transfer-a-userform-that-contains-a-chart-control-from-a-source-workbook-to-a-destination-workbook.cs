using System;
using System.IO;
using Aspose.Cells;

class TransferUserForm
{
    static void Main()
    {
        const string sourcePath = "SourceWithUserForm.xlsm";
        const string destinationPath = "DestinationWithUserForm.xlsm";

        // Verify that the source file exists to avoid FileNotFoundException.
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the source workbook (auto‑detects the .xlsm format).
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook.
            Workbook destinationWorkbook = new Workbook();

            // Copy the entire VBA project (including all UserForms) from source to destination.
            destinationWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

            // Save the destination workbook as a macro‑enabled file.
            destinationWorkbook.Save(destinationPath, SaveFormat.Xlsm);

            Console.WriteLine($"UserForm transferred successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}