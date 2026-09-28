// Title: Identify the encryption algorithm of an ODS workbook and display its security strength using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an ODS file with Aspose.Cells, catches the password‑required exception, opens the ODS ZIP package, reads the META‑INF/manifest.xml entry, locates the <manifest:algorithm> element, and prints a readable security‑level message. | Implement a helper that translates ODF algorithm‑name values such as AES256, AES128, Blowfish, and TripleDES into human‑friendly descriptions and integrates it into the Aspose.Cells loading flow.
// Common Searches: c# programmatically detect encryption algorithm used in an ODS workbook | aspocells read ODS zip manifest to find encryption details | how to map ODS algorithm-name attribute to security level in .NET | determine if ODS file is password protected using Aspose.Cells | extract encryption-data element from ODF manifest.xml with C#
// Tags: detect ODS encryption algorithm Aspose.Cells | parse ODS manifest.xml for encryption info | map ODF algorithm identifier to security level | verify ODS password protection using Aspose.Cells | extract encryption-data element from ODF manifest

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;

// Alias to avoid ambiguity with System.Xml.Linq.LoadOptions
using CellsLoadOptions = Aspose.Cells.LoadOptions;

// The example loads an ODS file with Aspose.Cells, detects password protection via exception handling, opens the ODS ZIP archive, reads META-INF/manifest.xml, extracts the <manifest:algorithm> element's algorithm-name attribute, maps known algorithms (AES256, AES128, Blowfish, TripleDES) to a human‑readable security description, and prints both the algorithm and its security level.
class OdsEncryptionDetector
{
    static void Main()
    {
        // Path to the ODS file
        string odsPath = "sample.ods";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(odsPath))
        {
            Console.WriteLine($"File not found: {odsPath}");
            return;
        }

        bool isEncrypted = false;
        Workbook workbook = null;

        // Attempt to load the workbook; if a password is required an exception will be thrown
        try
        {
            var loadOptions = new CellsLoadOptions(LoadFormat.Ods);
            workbook = new Workbook(odsPath, loadOptions);
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws a CellsException when a password is required or invalid
            if (!string.IsNullOrEmpty(ex.Message) &&
                (ex.Message.Contains("Password", StringComparison.OrdinalIgnoreCase)))
            {
                isEncrypted = true;
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        if (!isEncrypted)
        {
            Console.WriteLine("The ODS file is not encrypted.");
            return;
        }

        // Open the ODS (ZIP archive) to read the manifest
        try
        {
            using (ZipArchive zip = ZipFile.OpenRead(odsPath))
            {
                // The encryption information is stored in META-INF/manifest.xml
                ZipArchiveEntry manifestEntry = zip.GetEntry("META-INF/manifest.xml");
                if (manifestEntry == null)
                {
                    Console.WriteLine("Unable to locate manifest.xml; cannot determine encryption algorithm.");
                    return;
                }

                XDocument manifestDoc;
                using (var stream = manifestEntry.Open())
                {
                    manifestDoc = XDocument.Load(stream);
                }

                // Namespace used in ODF manifest files
                XNamespace ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

                // Find the <manifest:encryption-data> element inside any <manifest:file-entry>
                var algorithmElement = manifestDoc
                    .Descendants(ns + "encryption-data")
                    .Descendants(ns + "algorithm")
                    .FirstOrDefault();

                if (algorithmElement == null)
                {
                    Console.WriteLine("Encryption algorithm information not found in manifest.");
                    return;
                }

                // The algorithm name attribute (e.g., "AES256")
                string algorithmName = (string)algorithmElement.Attribute(ns + "algorithm-name") ?? "Unknown";

                // Map algorithm to a human‑readable security level description
                string description = algorithmName switch
                {
                    "AES256" => "AES‑256 (Strong encryption, considered very secure).",
                    "AES128" => "AES‑128 (Strong encryption, widely accepted as secure).",
                    "Blowfish" => "Blowfish (Moderate security; older algorithm).",
                    "TripleDES" => "Triple DES (Legacy encryption; weaker than modern standards).",
                    _ => $"Algorithm \"{algorithmName}\" detected (security level unknown)."
                };

                Console.WriteLine($"Encryption algorithm: {algorithmName}");
                Console.WriteLine($"Security level: {description}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing ODS archive: {ex.Message}");
        }
    }
}
