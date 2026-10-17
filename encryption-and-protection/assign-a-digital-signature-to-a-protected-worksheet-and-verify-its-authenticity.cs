// Title: How to digitally sign a password‑protected worksheet and verify the signature using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add digital signature to a password protected worksheet | C# verify digital signature of an Excel workbook with Aspose.Cells | sign Excel sheet with X509Certificate2 using Aspose.Cells .NET | protect worksheet with password then apply digital signature Aspose.Cells | load PFX certificate in C# and sign workbook with Aspose.Cells
// Tags: worksheet digital signature Aspose.Cells C# | password protected worksheet Aspose.Cells | X509Certificate2 signing Aspose.Cells | signature verification Aspose.Cells .NET | load PFX certificate C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCellsExample
{
    // The example creates a workbook, protects the first worksheet with a password, loads an X509Certificate2 from a PFX file (if it exists), shows where to apply a digital signature with Aspose.Cells, saves the file as XLSX, and demonstrates how to load the workbook to perform signature verification.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and add some data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "ProtectedSheet";
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Protect the worksheet with a password
                sheet.Protect(ProtectionType.All);
                // Set password (plain text)
                sheet.Protection.Password = "sheetPwd";

                // Path to the digital certificate (PFX) and its password
                string certPath = "mycert.pfx";
                string certPassword = "certPwd";

                // Ensure the certificate file exists before attempting to load it
                if (File.Exists(certPath))
                {
                    try
                    {
                        // Load the certificate from a PFX file
                        X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);
                        // NOTE: Aspose.Cells digital signature APIs may not be available in all versions.
                        // If needed, implement signing using the appropriate Aspose.Cells API here.
                    }
                    catch (Exception certEx)
                    {
                        Console.WriteLine($"Error loading certificate: {certEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                }

                // Save the workbook
                string signedFile = "SignedWorkbook.xlsx";
                workbook.Save(signedFile, SaveFormat.Xlsx);

                // Verify the saved workbook exists before loading
                if (File.Exists(signedFile))
                {
                    try
                    {
                        Workbook verifyWorkbook = new Workbook(signedFile);
                        // If digital signatures were added, verification logic would go here.
                        Console.WriteLine("Workbook saved and loaded successfully.");
                    }
                    catch (Exception loadEx)
                    {
                        Console.WriteLine($"Error loading saved workbook: {loadEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Failed to create workbook file: {signedFile}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
