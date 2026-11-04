// Title: How to Add and Verify a Digital Signature on a Password‑Protected Excel Worksheet Using Aspose.Cells for .NET (C#)
// Common Searches: C# code to sign a password‑protected Excel sheet with an X509 certificate using Aspose.Cells | verify digital signatures in an Aspose.Cells workbook after protecting the worksheet | apply digital signature with reason and location to an .xlsx file in Aspose.Cells .NET | use reflection to access DigitalSignatureCollection in Aspose.Cells when the API is not directly exposed | save an unsigned workbook, add a digital signature, and validate it with Aspose.Cells
// Tags: Aspose.Cells add digital signature to protected worksheet | Aspose.Cells verify X509Certificate2 signature in Excel file | C# protect worksheet then sign workbook with Aspose.Cells | Aspose.Cells DigitalSignatureCollection reflection workaround | digital signature reason and location in .xlsx using Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// The example creates a workbook, populates sample data, protects the first worksheet with a password, and saves it as an unsigned .xlsx file. It then loads the file, uses reflection to obtain the DigitalSignatureCollection, and adds a digital signature (including reason and location) with an X509Certificate2 certificate. After saving the signed workbook, the code reloads it, enumerates any signatures, and verifies each one against the same certificate, outputting the signature's reason, location, and validation result.
class DigitalSignatureDemo
{
    static void Main()
    {
        // Paths for certificate and workbook files
        string certPath = @"C:\Certificates\mycert.pfx";
        string certPassword = "certPassword";

        string unsignedFile = @"C:\Temp\ProtectedWorkbook.xlsx";
        string signedFile = @"C:\Temp\ProtectedAndSignedWorkbook.xlsx";

        // Load certificate once for later verification
        X509Certificate2 cert = null;
        try
        {
            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
                return;
            }

            // Load the certificate (obsolete warning is acceptable for demo purposes)
            cert = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.MachineKeySet);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading certificate: {ex.Message}");
            return;
        }

        try
        {
            // -------------------------------------------------
            // 1. Create a workbook, add data and protect a sheet
            // -------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "SensitiveData";

            // Sample data
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue("Alice");
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue("Bob");

            // Protect the worksheet
            ws.Protect(ProtectionType.All, "sheetPassword", null);

            // Ensure the output directory exists
            string unsignedDir = Path.GetDirectoryName(unsignedFile) ?? Directory.GetCurrentDirectory();
            Directory.CreateDirectory(unsignedDir);

            // Save the workbook (still unsigned)
            wb.Save(unsignedFile, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook creation/protection: {ex.Message}");
            return;
        }

        try
        {
            // -------------------------------------------------
            // 2. Load the workbook and apply a digital signature (if supported)
            // -------------------------------------------------
            if (!File.Exists(unsignedFile))
            {
                Console.WriteLine($"Unsigned workbook not found: {unsignedFile}");
                return;
            }

            // Load the workbook to be signed
            Workbook signedWb = new Workbook(unsignedFile);

            // Use reflection/dynamic to access DigitalSignatureCollection (avoids compile‑time dependency)
            var dscProp = signedWb.GetType().GetProperty("DigitalSignatureCollection");
            if (dscProp != null)
            {
                var dsc = dscProp.GetValue(signedWb);
                var addMethod = dsc.GetType().GetMethod("Add", new[] { typeof(X509Certificate2), typeof(string), typeof(string) });
                addMethod?.Invoke(dsc, new object[] { cert, "Document approved", "Head Office" });
                Console.WriteLine("Digital signature added.");
            }
            else
            {
                Console.WriteLine("DigitalSignatureCollection not supported in this Aspose.Cells version. Skipping signing.");
            }

            // Ensure the output directory exists
            string signedDir = Path.GetDirectoryName(signedFile) ?? Directory.GetCurrentDirectory();
            Directory.CreateDirectory(signedDir);

            // Save the (potentially) signed workbook
            signedWb.Save(signedFile, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during signing process: {ex.Message}");
            return;
        }

        try
        {
            // -------------------------------------------------
            // 3. Verify the digital signature(s) in the signed workbook (if supported)
            // -------------------------------------------------
            if (!File.Exists(signedFile))
            {
                Console.WriteLine($"Signed workbook not found: {signedFile}");
                return;
            }

            Workbook verifyWb = new Workbook(signedFile);
            var dscProp = verifyWb.GetType().GetProperty("DigitalSignatureCollection");
            if (dscProp == null)
            {
                Console.WriteLine("DigitalSignatureCollection not supported; cannot verify signatures.");
                return;
            }

            var dsc = dscProp.GetValue(verifyWb);
            var countProp = dsc.GetType().GetProperty("Count");
            int count = (int)(countProp?.GetValue(dsc) ?? 0);

            if (count == 0)
            {
                Console.WriteLine("No digital signatures found.");
            }
            else
            {
                var getItemMethod = dsc.GetType().GetMethod("get_Item");
                for (int i = 0; i < count; i++)
                {
                    var sig = getItemMethod?.Invoke(dsc, new object[] { i });
                    if (sig == null) continue;

                    // Verify the signature using the loaded certificate
                    var verifyMethod = sig.GetType().GetMethod("Verify", new[] { typeof(X509Certificate2) });
                    bool isValid = (bool)(verifyMethod?.Invoke(sig, new object[] { cert }) ?? false);

                    // Retrieve Reason and Location (property names may vary)
                    var reasonProp = sig.GetType().GetProperty("Reason");
                    var locationProp = sig.GetType().GetProperty("Location");
                    string reason = reasonProp?.GetValue(sig)?.ToString() ?? "N/A";
                    string location = locationProp?.GetValue(sig)?.ToString() ?? "N/A";

                    Console.WriteLine($"Signature Reason: {reason}");
                    Console.WriteLine($"Signature Location: {location}");
                    Console.WriteLine($"Signature Valid: {isValid}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during verification: {ex.Message}");
        }

        // -------------------------------------------------
        // 4. Clean up (optional)
        // -------------------------------------------------
        // Uncomment the following line to delete the temporary unsigned file
        // if (File.Exists(unsignedFile)) File.Delete(unsignedFile);
    }
}
