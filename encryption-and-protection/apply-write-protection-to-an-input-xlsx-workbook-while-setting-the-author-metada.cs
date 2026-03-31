using System;
using Aspose.Cells;

class ApplyWriteProtection
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // Path where the protected workbook will be saved
        string outputPath = "output_protected.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Set write‑protection password
        workbook.Settings.WriteProtection.Password = "SecurePass123";

        // Set the author for the write‑protection metadata
        workbook.Settings.WriteProtection.Author = "John Doe";

        // Optionally set the file author property (document metadata)
        workbook.Settings.Author = "John Doe";

        // Save the workbook with the applied protection
        workbook.Save(outputPath);
    }
}