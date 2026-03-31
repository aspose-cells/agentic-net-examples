using System;
using System.IO;
using Aspose.Cells;

class TransferPageSetup
{
    static void Main()
    {
        // Load the source workbook that contains the desired page setup configuration
        string sourcePath = "source.xlsx";
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException($"Source file not found: {sourcePath}");
        }
        Workbook sourceWorkbook = new Workbook(sourcePath);

        // Load the target workbook where the page setup will be applied, or create a new one if it doesn't exist
        string targetPath = "target.xlsx";
        Workbook targetWorkbook;
        if (File.Exists(targetPath))
        {
            targetWorkbook = new Workbook(targetPath);
        }
        else
        {
            targetWorkbook = new Workbook();
        }

        // Access the source worksheet (adjust index or name as needed)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Access the target worksheet (adjust index or name as needed)
        Worksheet targetSheet = targetWorkbook.Worksheets[0];

        // Transfer the page setup from source to target using the PageSetup.Copy method
        targetSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

        // Save the target workbook with the updated page setup
        targetWorkbook.Save("target_updated.xlsx");
    }
}