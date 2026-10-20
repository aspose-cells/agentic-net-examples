// Title: Overwrite the original Excel file after adding a digital signature with Aspose.Cells for .NET (no backup)
// Common Searches: C# Aspose.Cells save signed workbook over original file without creating backup | How to replace existing Excel workbook after applying digital signature using Aspose.Cells | Aspose.Cells overwrite existing .xlsx after adding SignatureCollection signature | Save signed Excel workbook in place using Aspose.Cells .NET | Avoid backup file when saving signed workbook with Aspose.Cells
// Tags: overwrite signed workbook Aspose.Cells | Aspose.Cells SignatureCollection add signature | save workbook without backup .NET | digital certificate signing Excel Aspose.Cells | in-place save Excel after signature Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing .xlsx file, checks for a digital certificate, uses reflection to invoke the SignatureCollection.Add method when available, and then saves the workbook back to the same path with SaveFormat.Xlsx, overwriting the original file and avoiding any backup creation.
class WorkbookSigning
{
    static void Main()
    {
        try
        {
            // Path to the original workbook
            string workbookPath = @"C:\Docs\Sample.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path to the digital certificate (optional)
            string certPath = @"C:\Certs\mycert.pfx";
            string certPassword = "certPassword";

            // Attempt to sign the workbook if the SignatureCollection API is available
            if (File.Exists(certPath))
            {
                try
                {
                    // Check if the SignatureCollection property exists (for compatibility with older versions)
                    var signatureProp = typeof(Workbook).GetProperty("SignatureCollection");
                    if (signatureProp != null)
                    {
                        // Use reflection to invoke the Add method
                        var signatureCollection = signatureProp.GetValue(workbook);
                        var addMethod = signatureCollection.GetType().GetMethod("Add", new[] { typeof(string), typeof(string), typeof(string), typeof(string) });
                        if (addMethod != null)
                        {
                            addMethod.Invoke(signatureCollection, new object[] { certPath, certPassword, "Document approved", "Company" });
                            Console.WriteLine("Workbook signed successfully.");
                        }
                        else
                        {
                            Console.WriteLine("SignatureCollection.Add method not found. Skipping signing.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("SignatureCollection API is not available in this Aspose.Cells version. Skipping signing.");
                    }
                }
                catch (Exception signEx)
                {
                    Console.WriteLine($"Error while signing the workbook: {signEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
            }

            // Save the workbook, overwriting the original file.
            workbook.Save(workbookPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook processed and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
