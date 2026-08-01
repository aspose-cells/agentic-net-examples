// Title: C# – Wrap Aspose.Cells Digital Signature in Try‑Catch and Log Errors
// Description: Shows how to create a Workbook, add content, load an X509Certificate2, build a DigitalSignature, and apply it with Workbook.SetDigitalSignature inside a try‑catch block that records the exception type, message and stack trace before saving the file.
// Keywords: Aspose.Cells digital signature C# | try catch Aspose.Cells | exception handling Aspose.Cells | log digital signature errors .NET | Workbook.SetDigitalSignature error handling | X509Certificate2 signing Aspose | C# workbook signing example | Aspose.Cells error logging | digital signature exception logging | C# try‑catch logging pattern
// Common Searches: how to catch errors when using Aspose.Cells SetDigitalSignature in C# | Aspose.Cells digital signature exception handling example | log certificate signing failures with Aspose.Cells | C# try‑catch around Aspose.Cells digital signature | Aspose.Cells SetDigitalSignature stack trace logging
// Developer Intent: Add robust try‑catch handling around the Aspose.Cells digital signing process and output detailed error information for troubleshooting.
// Use Cases: Prevent application crashes if the certificate file is missing, corrupted, or the password is wrong. | Capture diagnostic data (exception type, message, stack trace) to speed up support and debugging. | Allow the workbook to be saved even when the signing step fails, preserving unsignaled content. | Integrate with existing logging frameworks (e.g., NLog, Serilog) by replacing console writes.
// AI Prompts: Generate C# code that wraps Aspose.Cells SetDigitalSignature in a try‑catch block and logs exception details using a logging framework. | Suggest best practices for handling certificate loading errors and digital signature failures in Aspose.Cells. | Create a reusable method that signs a workbook with Aspose.Cells and returns a result object containing success status and error information.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Shows how to create a Workbook, add content, load an X509Certificate2, build a DigitalSignature, and apply it with Workbook.SetDigitalSignature inside a try‑catch block that records the exception type, message and stack trace before saving the file.
class DigitalSignatureDemo
{
    static void Main()
    {
        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Digitally Signed Document");

        try
        {
            // Load the certificate (replace with actual path and password)
            X509Certificate2 certificate = new X509Certificate2("mycert.pfx", "password");

            // Create a digital signature using the certificate, comments, and sign time
            DigitalSignature signature = new DigitalSignature(
                certificate,
                "Document approval",
                DateTime.UtcNow);

            // Add the signature to a collection
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);

            // Apply the digital signature collection to the workbook
            workbook.SetDigitalSignature(signatures);
        }
        catch (Exception ex)
        {
            // Log detailed error information
            Console.WriteLine($"Error signing workbook: {ex.GetType().Name} - {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }

        // Save the workbook (signed if no exception occurred)
        workbook.Save("signed_output.xlsx");
    }
}
