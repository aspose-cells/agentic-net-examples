// Title: C# – Detect Encrypted .xlsx and Retrieve Encryption Algorithm URI with Aspose.Cells
// Description: A concise example that uses Aspose.Cells FileFormatUtil to check if an Excel workbook is encrypted, then opens the ZIP package, reads the EncryptionInfo part, and extracts the algorithm URI and key size for compliance auditing.
// Keywords: Aspose.Cells encryption detection | C# read EncryptionInfo | encrypted xlsx algorithm URI | Office Open XML encryption audit | FileFormatUtil IsEncrypted | extract key size .NET | Excel file security compliance
// Common Searches: How to detect encrypted Excel file with Aspose.Cells C# | Get encryption algorithm URI from encrypted .xlsx | Read EncryptionInfo part of Office Open XML | Audit Excel workbook encryption algorithm | C# extract key size from encrypted workbook
// Developer Intent: Determine whether an .xlsx file is encrypted and obtain its encryption algorithm identifier and key length for security auditing.
// Use Cases: Batch‑scan Excel files before archiving to ensure only approved encryption algorithms are used. | Log algorithm URI and key size of incoming workbooks for regulatory reporting. | Validate encrypted workbooks against corporate security policies during automated import processes.
// AI Prompts: Write C# code that uses Aspose.Cells to detect encryption and extracts the algorithm URI and key size from the EncryptionInfo part of an .xlsx file. | Provide robust error handling for ZIP archive processing when the file is not a valid Office Open XML package. | Show how to extend this example to scan a directory of workbooks and output results to a CSV file.

using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells;

namespace AsposeCellsEncryptionAudit
{
    // A concise example that uses Aspose.Cells FileFormatUtil to check if an Excel workbook is encrypted, then opens the ZIP package, reads the EncryptionInfo part, and extracts the algorithm URI and key size for compliance auditing.
    class Program
    {
        static void Main()
        {
            // Path to the Office Open XML file to be audited
            string filePath = "encrypted.xlsx";

            // Ensure the file exists before proceeding
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Detect file format and encryption status using Aspose.Cells
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"Is the file encrypted? {formatInfo.IsEncrypted}");

                if (formatInfo.IsEncrypted)
                {
                    // Aspose.Cells does not expose detailed encryption info directly,
                    // so we fall back to manual extraction of the EncryptionInfo part.
                    Console.WriteLine("Attempting manual extraction of EncryptionInfo part...");

                    try
                    {
                        using (FileStream fs = File.OpenRead(filePath))
                        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                        {
                            ZipArchiveEntry encryptionInfoEntry = zip.GetEntry("EncryptionInfo");
                            if (encryptionInfoEntry != null)
                            {
                                using (Stream entryStream = encryptionInfoEntry.Open())
                                {
                                    XDocument encryptionDoc = XDocument.Load(entryStream);
                                    XNamespace encNs = "http://schemas.microsoft.com/office/2006/encryption";

                                    XElement algorithmElement = encryptionDoc.Root?.Element(encNs + "algorithm");
                                    if (algorithmElement != null)
                                    {
                                        string algorithmUri = (string)algorithmElement.Attribute("uri");
                                        string keySize = (string)algorithmElement.Attribute("keySize");

                                        Console.WriteLine($"Encryption algorithm URI: {algorithmUri}");
                                        Console.WriteLine($"Key size (bits): {keySize}");
                                    }
                                    else
                                    {
                                        Console.WriteLine("Algorithm element not found in EncryptionInfo.");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("EncryptionInfo part not found; unable to determine algorithm.");
                            }
                        }
                    }
                    catch (InvalidDataException)
                    {
                        Console.WriteLine("The file is encrypted and cannot be treated as a standard ZIP archive.");
                    }
                    catch (Exception zipEx)
                    {
                        Console.WriteLine($"Error while processing ZIP archive: {zipEx.Message}");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred during processing: {e.Message}");
            }
        }
    }
}
