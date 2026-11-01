// Title: Export a VBA project's digital certificate to a MemoryStream using Aspose.Cells for .NET
// Common Searches: how to extract VBA macro digital certificate to a MemoryStream using Aspose.Cells C# | Aspose.Cells export VBA project digital signature as byte array | save VBA project certificate from an Excel workbook to a .cer file with Aspose.Cells | retrieve VBA project DigitalSignature property in Aspose.Cells .NET | C# code to export VBA digital certificate from .xlsx using Aspose.Cells
// Tags: export VBA digital certificate Aspose.Cells | retrieve VBA project DigitalSignature .NET | memory stream certificate extraction C# | save VBA certificate to .cer file | Aspose.Cells VBA macro certificate handling

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, checks for a VBA project, uses reflection to read the VbaProject's DigitalSignature byte array, writes the bytes into a MemoryStream, resets the stream, and optionally saves the certificate as a .cer file.
class ExportVbaCertificate
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Prepare a memory stream to hold the exported certificate
        using (MemoryStream certificateStream = new MemoryStream())
        {
            try
            {
                // Check if the workbook contains a VBA project
                var vbaProject = workbook.VbaProject;
                if (vbaProject != null)
                {
                    // Use reflection to obtain the DigitalSignature property (may not exist in older versions)
                    var digitalSignatureProp = vbaProject.GetType().GetProperty("DigitalSignature");
                    if (digitalSignatureProp != null)
                    {
                        var certBytes = digitalSignatureProp.GetValue(vbaProject) as byte[];
                        if (certBytes != null && certBytes.Length > 0)
                        {
                            // Write the certificate bytes into the memory stream
                            certificateStream.Write(certBytes, 0, certBytes.Length);
                            certificateStream.Position = 0; // Reset for reading

                            // Example: save the certificate to a file
                            const string outputPath = "VbaCertificate.cer";
                            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                            {
                                certificateStream.CopyTo(file);
                            }

                            Console.WriteLine($"VBA digital certificate exported to \"{outputPath}\".");
                        }
                        else
                        {
                            Console.WriteLine("No digital certificate data found in the VBA project.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("The VbaProject does not expose a DigitalSignature property in this Aspose.Cells version.");
                    }
                }
                else
                {
                    Console.WriteLine("No VBA project found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while exporting the certificate: {ex.Message}");
            }
        }
    }
}
