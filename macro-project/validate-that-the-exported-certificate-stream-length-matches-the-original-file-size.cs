// Title: Check that an exported XLSX MemoryStream size matches the original file size using Aspose.Cells for .NET
// Common Searches: how to verify that a MemoryStream saved with Aspose.Cells has the same byte length as the source XLSX file | C# compare original Excel file size with exported workbook stream size using Aspose.Cells | Aspose.Cells .NET validate exported workbook size equals input file size | detect size mismatch after saving workbook to MemoryStream with Aspose.Cells | handle password protected workbook when checking exported stream length in C#
// Tags: Aspose.Cells workbook export MemoryStream | validate exported XLSX byte length | original vs exported Excel file size comparison | password protected workbook loading Aspose.Cells | file existence check C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX file with Aspose.Cells, obtains its byte length, saves the workbook to a MemoryStream using SaveFormat.Xlsx, and compares the stream length to the original file size, reporting success or mismatch while handling missing files, password‑protected workbooks, and unexpected errors.
class CertificateExportValidator
{
    static void Main()
    {
        // Path to the original Excel file
        string filePath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found – \"{filePath}\"");
            return;
        }

        try
        {
            // Get original file size in bytes
            long originalFileSize = new FileInfo(filePath).Length;

            // Load the workbook (handle possible password protection)
            Workbook workbook;
            try
            {
                workbook = new Workbook(filePath);
            }
            catch (CellsException ex) when (ex.Message.Contains("password"))
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password to load it.");
                return;
            }

            // Export the workbook to a memory stream in the same format
            using (MemoryStream exportStream = new MemoryStream())
            {
                workbook.Save(exportStream, SaveFormat.Xlsx);

                // Length of the exported stream
                long exportedStreamLength = exportStream.Length;

                // Validate that the exported stream length matches the original file size
                if (exportedStreamLength == originalFileSize)
                {
                    Console.WriteLine("Validation succeeded: exported stream length matches the original file size.");
                }
                else
                {
                    Console.WriteLine($"Validation failed: original size = {originalFileSize} bytes, exported size = {exportedStreamLength} bytes.");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
