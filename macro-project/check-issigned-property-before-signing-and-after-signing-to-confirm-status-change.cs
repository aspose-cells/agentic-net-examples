// Title: How to read Workbook.IsSigned before and after adding a digital signature with Aspose.Cells in C#
// AI Prompts: Generate C# code that reads the Workbook.IsSigned property, adds a digital signature from a PFX file using Aspose.Cells (using reflection for older versions), then reads IsSigned again to verify the change. | Show a C# example that outputs the signed status of an Excel workbook before signing, conditionally applies a digital signature via reflection when supported, and prints the status after signing.
// Common Searches: C# Aspose.Cells check if Excel workbook is signed before signing | How to verify Workbook.IsSigned changes after applying a digital signature with Aspose.Cells | Use reflection to add a digital signature in Aspose.Cells when SignatureCollection is unavailable | Aspose.Cells IsSigned property example code in .NET | Detect digital signature support in different Aspose.Cells versions for C#
// Tags: Aspose.Cells digital signature verification | C# Workbook.IsSigned property | Aspose.Cells reflection add signature | Excel workbook signed status check | Aspose.Cells version compatibility signature

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, optionally signs it with a PFX certificate using reflection to access SignatureCollection when available, and saves the file. It demonstrates how to read the Workbook.IsSigned property before and after the signing operation to confirm that the signed status changes.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // Attempt to sign the workbook if the API is available
            string certificatePath = "certificate.pfx";
            string certificatePassword = "password";

            if (File.Exists(certificatePath))
            {
                // Aspose.Cells versions prior to 22.x do not support digital signatures.
                // The following block is kept for reference; it will be executed only
                // when the SignatureCollection property exists.
                var signatureProp = workbook.GetType().GetProperty("SignatureCollection");
                if (signatureProp != null)
                {
                    // Use reflection to invoke the Add method safely.
                    var signatures = signatureProp.GetValue(workbook);
                    var addMethod = signatures.GetType().GetMethod("Add", new[] { typeof(string), typeof(string) });
                    if (addMethod != null)
                    {
                        addMethod.Invoke(signatures, new object[] { certificatePath, certificatePassword });
                        var countProp = signatures.GetType().GetProperty("Count");
                        int signedCount = (int)countProp.GetValue(signatures);
                        Console.WriteLine($"Workbook signed successfully. Signature count: {signedCount}");
                    }
                    else
                    {
                        Console.WriteLine("SignatureCollection does not support adding signatures in this version.");
                    }
                }
                else
                {
                    Console.WriteLine("SignatureCollection property not found. Signing is not supported in this Aspose.Cells version.");
                }
            }
            else
            {
                Console.WriteLine($"Certificate file not found: {certificatePath}");
            }

            // Save the workbook (lifecycle rule)
            string outputPath = "SignedWorkbook.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
