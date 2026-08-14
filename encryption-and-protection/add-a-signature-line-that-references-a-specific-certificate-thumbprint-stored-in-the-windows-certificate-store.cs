// Title: Add a Signature Line and Digitally Sign an Excel Workbook Using a Certificate Thumbprint (Aspose.Cells for .NET)
// Description: Loads an X509Certificate2 from the CurrentUser\My store by thumbprint, creates a SignatureLine with signer details, links it to a DigitalSignature, adds the signature to the workbook’s DigitalSignatureCollection, and saves the signed Excel file.
// Keywords: Aspose.Cells signature line | digital signature certificate thumbprint | C# Excel signing | X509Store Aspose.Cells | add signature line .NET | certificate thumbprint Excel | Windows certificate store signing
// Common Searches: Aspose.Cells add signature line and sign workbook | C# load certificate by thumbprint for Excel digital signature | How to use X509Store with Aspose.Cells | Link SignatureLine to DigitalSignature in Aspose.Cells | Apply digital signature to Excel file using Windows certificate
// Developer Intent: Create a pre‑filled signature line in an Excel workbook and apply a digital signature using a certificate identified by its thumbprint from the Windows certificate store.
// Use Cases: Automated approval sheets where the signature line is pre‑populated and the file is signed with a corporate certificate retrieved by thumbprint. | Compliance reporting that programmatically adds a signature line to each generated report and signs it without user interaction. | Graceful handling of missing or invalid thumbprints by skipping the digital signature while still producing a usable workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds a SignatureLine to a worksheet and signs the workbook using a certificate loaded from the Windows certificate store by thumbprint. | Explain how to detect and handle a missing certificate thumbprint when applying a digital signature with Aspose.Cells. | Show how to associate a DigitalSignature Id with an existing SignatureLine to ensure the signature matches the correct line in Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureDemo
{
    // Loads an X509Certificate2 from the CurrentUser\My store by thumbprint, creates a SignatureLine with signer details, links it to a DigitalSignature, adds the signature to the workbook’s DigitalSignatureCollection, and saves the signed Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Thumbprint of the certificate stored in the Windows certificate store
                const string certificateThumbprint = "YOUR_CERTIFICATE_THUMBPRINT_HERE";

                X509Certificate2 certificate = null;
                if (!string.IsNullOrWhiteSpace(certificateThumbprint) &&
                    !certificateThumbprint.Equals("YOUR_CERTIFICATE_THUMBPRINT_HERE", StringComparison.OrdinalIgnoreCase))
                {
                    // Load the certificate from CurrentUser\My store
                    using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
                    {
                        store.Open(OpenFlags.ReadOnly);
                        X509Certificate2Collection found = store.Certificates.Find(
                            X509FindType.FindByThumbprint,
                            certificateThumbprint,
                            validOnly: false);

                        if (found.Count > 0)
                        {
                            certificate = found[0];
                        }
                        else
                        {
                            Console.WriteLine($"Certificate with thumbprint {certificateThumbprint} not found. Continuing without digital signature.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No valid certificate thumbprint provided. Skipping digital signature.");
                }

                // Create a new workbook
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Create a signature line and assign an identifier
                SignatureLine sigLine = new SignatureLine
                {
                    Signer = "John Doe",
                    Title = "Approver",
                    Email = "john.doe@example.com",
                    IsLine = true,
                    AllowComments = true,
                    ShowSignedDate = true,
                    Instructions = "Please sign to approve.",
                    Id = Guid.NewGuid()
                };

                // Add the signature line to the worksheet (row 5, column 2 as an example)
                ws.Shapes.AddSignatureLine(5, 2, sigLine);

                // If a certificate was loaded, create and attach a digital signature
                if (certificate != null)
                {
                    DigitalSignature digitalSignature = new DigitalSignature(
                        certificate,
                        "Approved by John Doe",
                        DateTime.Now)
                    {
                        Id = sigLine.Id // Link to the signature line
                    };

                    DigitalSignatureCollection dsCollection = new DigitalSignatureCollection();
                    dsCollection.Add(digitalSignature);
                    wb.SetDigitalSignature(dsCollection);
                }

                // Save the signed workbook
                const string outputPath = "SignedWorkbook_WithSignatureLine.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
