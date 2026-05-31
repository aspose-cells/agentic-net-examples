using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "copied_protected.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Access the first worksheet in each workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define the source range to copy (e.g., cells A1:B5)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:B5");

            // Define the destination range where the data will be pasted (e.g., cells C1:D5)
            Aspose.Cells.Range destinationRange = destinationSheet.Cells.CreateRange("C1:D5");

            // Copy the source range to the destination range
            destinationRange.Copy(sourceRange);

            // Protect the workbook with a password (file encryption)
            destinationWorkbook.Settings.Password = "MySecretPassword";

            // Save the new workbook to disk
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Workbook saved successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}