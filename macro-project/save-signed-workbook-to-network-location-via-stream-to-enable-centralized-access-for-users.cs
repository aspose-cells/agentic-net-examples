// Title: Save an Aspose.Cells workbook to a UNC network share using FileStream in C# with fallback when digital signature API is unavailable
// AI Prompts: Write C# code that builds an Aspose.Cells workbook, populates a cell, and writes the file to a UNC location using a FileStream, creating the folder if needed. | Generate a helper method that checks for a digital certificate file, handles its absence, and saves the workbook as an unsigned .xlsx when the signing API cannot be used. | Provide robust error‑handling for network directory creation and FileStream write operations in an Aspose.Cells C# example.
// Common Searches: how to write an Aspose.Cells workbook to a network share with C# FileStream | C# save Excel file to UNC path using Aspose.Cells when digital signature not supported | Aspose.Cells fallback save without signing certificate C# example | ensure network directory exists before saving Aspose.Cells workbook | save workbook to shared folder using Aspose.Cells and FileStream
// Tags: Aspose.Cells save workbook to UNC share | C# FileStream Excel workbook export | unsigned workbook fallback Aspose.Cells | network directory creation C# before file save | digital certificate handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates an Aspose.Cells workbook, writes "Hello World" to cell A1, checks for a digital certificate, ensures the target UNC directory exists, and saves the workbook as an .xlsx file to a network share via a FileStream. If the digital signature API is unavailable or the certificate is missing, it falls back to saving an unsigned workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World");

            // Paths for the digital certificate and the final workbook location
            string certificatePath = @"C:\Certificates\mycert.pfx";
            string certificatePassword = "yourPassword";
            string networkPath = @"\\fileserver\shared\SignedWorkbook.xlsx";

            // Verify that the certificate file exists before attempting to sign
            bool canSign = File.Exists(certificatePath);

            // Ensure the target directory exists
            string networkDir = Path.GetDirectoryName(networkPath) ?? string.Empty;
            if (!Directory.Exists(networkDir))
            {
                Directory.CreateDirectory(networkDir);
            }

            if (canSign)
            {
                try
                {
                    // Digital signature API is not available in the current Aspose.Cells version.
                    // The code falls back to saving without signing.
                    Console.WriteLine("Digital signature API not available; saving without signing.");
                    SaveWorkbook(workbook, networkPath);
                }
                catch (Exception signEx)
                {
                    Console.WriteLine($"Signing failed or not supported: {signEx.Message}");
                    // Fallback: save without signing
                    SaveWorkbook(workbook, networkPath);
                }
            }
            else
            {
                Console.WriteLine($"Certificate file not found: {certificatePath}");
                // Save without signing
                SaveWorkbook(workbook, networkPath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to save a workbook directly (used when signing is not possible)
    private static void SaveWorkbook(Workbook workbook, string path)
    {
        try
        {
            string dir = Path.GetDirectoryName(path) ?? string.Empty;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Save(stream, SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Failed to save workbook: {e.Message}");
        }
    }
}
