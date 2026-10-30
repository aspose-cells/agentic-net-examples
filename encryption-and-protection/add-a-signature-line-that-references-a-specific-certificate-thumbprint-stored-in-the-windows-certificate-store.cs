// Title: How to add a signature line to an Excel workbook in C# with Aspose.Cells using a Windows certificate thumbprint
// Common Searches: aspnet add signature line to Excel file using Aspose.Cells and certificate thumbprint | c# retrieve X509Certificate2 from Windows store to sign an Excel workbook with Aspose.Cells | using reflection to add SignatureLineCollection in older Aspose.Cells versions | sign Excel workbook programmatically with a specific certificate thumbprint in .NET | handle missing SignatureLineCollection API when adding a signature line with Aspose.Cells
// Tags: Aspose.Cells signature line API usage | C# digital signing of Excel files | X509Certificate2 lookup by thumbprint in .NET | reflection fallback for Aspose.Cells features | Windows certificate store integration with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Security.Cryptography.X509Certificates;

// The example creates a new Workbook, uses reflection to access Worksheet.SignatureLineCollection when available, adds a signature line, sets suggested signer name and email, retrieves an X509Certificate2 from the current user's Windows certificate store by thumbprint, signs the line with that certificate, and saves the workbook as SignedWorkbook.xlsx, handling missing API or certificate errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Add a signature line (available in newer Aspose.Cells versions)
            // ------------------------------------------------------------
            // The SignatureLineCollection API may not be present in older versions.
            // This block is wrapped in a separate try-catch and uses reflection to
            // avoid compile‑time errors if the property is missing.
            try
            {
                var signatureLineProp = typeof(Worksheet).GetProperty("SignatureLineCollection");
                if (signatureLineProp != null)
                {
                    var collection = signatureLineProp.GetValue(worksheet);
                    var addMethod = collection?.GetType().GetMethod("Add", new[] { typeof(int), typeof(int), typeof(double), typeof(double) });
                    if (addMethod != null)
                    {
                        // Add signature line at row 1, column 1 with width 200 and height 50
                        var signatureLine = addMethod.Invoke(collection, new object[] { 1, 1, 200.0, 50.0 });
                        var propertiesProp = signatureLine?.GetType().GetProperty("SignatureLineProperties");
                        var properties = propertiesProp?.GetValue(signatureLine);
                        var suggestedSignerProp = properties?.GetType().GetProperty("SuggestedSigner");
                        suggestedSignerProp?.SetValue(properties, "John Doe");
                        var suggestedEmailProp = properties?.GetType().GetProperty("SuggestedSignerEmail");
                        suggestedEmailProp?.SetValue(properties, "john.doe@example.com");

                        const string thumbprint = "YOUR_CERTIFICATE_THUMBPRINT";
                        X509Certificate2 cert = GetCertificateByThumbprint(thumbprint);
                        if (cert != null)
                        {
                            var signMethod = signatureLine?.GetType().GetMethod("Sign", new[] { typeof(X509Certificate2) });
                            signMethod?.Invoke(signatureLine, new object[] { cert });
                        }
                        else
                        {
                            Console.WriteLine("Certificate with the specified thumbprint was not found.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("SignatureLineCollection API is not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Signature line could not be added: {ex.Message}");
            }

            // Save the workbook
            string outputPath = "SignedWorkbook.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to locate a certificate by thumbprint
    static X509Certificate2 GetCertificateByThumbprint(string thumbprint)
    {
        try
        {
            // Normalize thumbprint (remove spaces, make uppercase)
            string normalizedThumbprint = thumbprint.Replace(" ", string.Empty).ToUpperInvariant();

            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                foreach (var cert in store.Certificates)
                {
                    if (!string.IsNullOrEmpty(cert.Thumbprint) &&
                        cert.Thumbprint.Equals(normalizedThumbprint, StringComparison.OrdinalIgnoreCase))
                    {
                        return cert;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing certificate store: {ex.Message}");
        }

        return null;
    }
}
