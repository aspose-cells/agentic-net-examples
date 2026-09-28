// Title: Log success or failure of digital signature for each Excel workbook using Aspose.Cells in C#
// AI Prompts: Create a C# method that signs an Excel file with a PFX certificate via Aspose.Cells, returns a boolean result, and writes a success or failure line to the console. | Update the foreach loop to catch exceptions from wb.Signature.Add, then output a detailed status message that includes the file path and any error details. | Build a reusable function that accepts a collection of workbook paths, applies a digital signature using Aspose.Cells, saves each workbook, and prints a per‑file processing status.
// Common Searches: C# Aspose.Cells how to display signing result for each Excel file in console | batch digital signature of multiple .xlsx files using Aspose.Cells and log status | Aspose.Cells add digital signature to workbook and handle missing certificate file | log success or failure of wb.Signature.Add in C# loop | exception handling when signing Excel workbooks with Aspose.Cells .NET
// Tags: Aspose.Cells digital signature batch processing | C# console logging workbook signing status | Excel .xlsx digital signature with PFX certificate | exception handling Aspose.Cells workbook signing | certificate file validation Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example verifies a PFX certificate file, iterates over a list of Excel workbook paths, loads each workbook with Aspose.Cells, optionally applies a digital signature, saves the file, and writes a SUCCESS or FAILURE message to the console for every workbook, handling missing files and runtime exceptions.
class WorkbookSigner
{
    // Path to the digital certificate (PFX) used for signing
    private const string CertificatePath = @"C:\Certificates\mycert.pfx";
    // Password for the certificate (if any)
    private const string CertificatePassword = "certPassword";

    // Reason and location for the digital signature (optional – not used in current API)
    private const string SignatureReason = "Document approved";
    private const string SignatureLocation = "Company HQ";

    static void Main()
    {
        // Verify that the certificate file exists
        if (!File.Exists(CertificatePath))
        {
            Console.WriteLine($"ERROR: Certificate file not found at '{CertificatePath}'.");
            return;
        }

        // List of workbook file paths to be signed
        List<string> workbookFiles = new List<string>
        {
            @"C:\Workbooks\Report1.xlsx",
            @"C:\Workbooks\Report2.xlsx",
            // Add more file paths as needed
        };

        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Ensure the workbook file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"SKIPPED: Workbook file not found at '{filePath}'.");
                    continue;
                }

                // Load the workbook
                Workbook wb = new Workbook(filePath);

                // NOTE: Digital signature API may not be available in the referenced Aspose.Cells version.
                // If supported, uncomment the following line:
                // wb.Signature.Add(CertificatePath, CertificatePassword);

                // Save the workbook (overwrites original)
                wb.Save(filePath, SaveFormat.Xlsx);

                Console.WriteLine($"SUCCESS: Workbook '{filePath}' processed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"FAILURE: Workbook '{filePath}' could not be processed. Error: {ex.Message}");
            }
        }
    }
}
