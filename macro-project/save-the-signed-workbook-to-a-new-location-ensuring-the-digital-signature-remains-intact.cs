// Title: Copy a digitally signed Excel workbook to another folder while preserving its signature using Aspose.Cells for .NET
// AI Prompts: Load a signed .xlsx workbook with Aspose.Cells and save it to a different path while keeping the existing digital signature unchanged. | Write C# code that duplicates a signed Excel file to a new location, ensuring the signature remains intact after the save operation.
// Common Searches: Aspose.Cells keep digital signature when saving a copied workbook in C# | How to copy a signed Excel file to another directory without losing the signature using .NET | Saving signed .xlsx with Aspose.Cells without removing the digital signature | C# Aspose.Cells duplicate signed workbook maintain signature
// Tags: preserve digital signature Aspose.Cells | copy signed Excel workbook C# | save signed workbook without modifying signature | Aspose.Cells SaveFormat.Xlsx preserving signature | load signed .xlsx with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a digitally signed Excel workbook, creates the target directory if needed, and saves a copy to a new location using Aspose.Cells for .NET, ensuring that the original digital signatures are retained.
class PreserveSignature
{
    static void Main()
    {
        string sourcePath = @"C:\Source\SignedWorkbook.xlsx";
        string destinationPath = @"C:\Destination\SignedWorkbook_Copy.xlsx";

        // Verify that the source workbook exists before attempting to load it.
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the signed workbook from its original location.
            Workbook workbook = new Workbook(sourcePath);

            // Ensure the destination directory exists.
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the workbook to a new location.
            // The digital signature(s) are preserved because the workbook is not modified.
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook copied successfully to: {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
