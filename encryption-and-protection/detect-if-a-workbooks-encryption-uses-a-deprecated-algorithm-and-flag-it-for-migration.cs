// Title: Check for deprecated encryption algorithms in an Excel workbook with Aspose.Cells for .NET and flag for migration
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, detects if it requires a password, and uses reflection to read the Workbook.EncryptionInfo.Algorithm property. | Create a method that evaluates the retrieved algorithm name, logs a warning when it is Xor, RC4, or RC4CryptoAPI, and suggests re‑encrypting the file using AES‑256.
// Common Searches: aspnet how to determine encryption algorithm of a password protected Excel file using Aspose.Cells | c# detect deprecated RC4 or XOR encryption in Excel workbook with Aspose.Cells | retrieve EncryptionInfo.Algorithm property via reflection Aspose.Cells .NET | flag Excel files using old encryption algorithms for migration to AES256 in C# | check if Excel workbook is encrypted and get algorithm name using Aspose.Cells LoadOptions
// Tags: identify legacy encryption algorithm Aspose.Cells | read EncryptionInfo.Algorithm via reflection C# | mark workbook with obsolete encryption | recommend AES256 re‑encryption for old Excel files | detect password protected workbook Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace WorkbookEncryptionChecker
{
    // The sample loads a specified .xlsx file with Aspose.Cells, determines whether it is password‑protected, uses reflection to access the Workbook.EncryptionInfo.Algorithm property, reports the detected algorithm, and issues a warning for deprecated algorithms (Xor, RC4, RC4CryptoAPI) while recommending migration to AES‑256 encryption.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string workbookPath = @"C:\Path\To\Your\Workbook.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            try
            {
                // Attempt to load the workbook without a password
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(workbookPath, loadOptions);

                // If loading succeeds, the workbook is not encrypted
                Console.WriteLine("The workbook is not encrypted.");
                return;
            }
            catch (CellsException ex) when (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
            {
                // The exception indicates that a password is required → workbook is encrypted
                Console.WriteLine("The workbook is encrypted (password protected).");

                // Try to obtain encryption details via reflection (avoids compile‑time dependency on specific API versions)
                try
                {
                    // Load the workbook again, this time allowing Aspose.Cells to expose encryption info
                    LoadOptions loadOptions = new LoadOptions { Password = "" };
                    Workbook encryptedWorkbook = new Workbook(workbookPath, loadOptions);

                    // Use reflection to get the EncryptionInfo property
                    PropertyInfo encInfoProp = typeof(Workbook).GetProperty("EncryptionInfo", BindingFlags.Public | BindingFlags.Instance);
                    if (encInfoProp != null)
                    {
                        object encInfo = encInfoProp.GetValue(encryptedWorkbook);
                        if (encInfo != null)
                        {
                            // Get the Algorithm property from the EncryptionInfo object
                            PropertyInfo algorithmProp = encInfo.GetType().GetProperty("Algorithm", BindingFlags.Public | BindingFlags.Instance);
                            object algorithmValue = algorithmProp?.GetValue(encInfo);
                            string algorithmName = algorithmValue?.ToString() ?? "Unknown";

                            Console.WriteLine($"Encryption algorithm detected: {algorithmName}");

                            // Determine if the algorithm is deprecated (based on name)
                            bool isDeprecated = algorithmName.Equals("Xor", StringComparison.OrdinalIgnoreCase) ||
                                                algorithmName.Equals("RC4", StringComparison.OrdinalIgnoreCase) ||
                                                algorithmName.Equals("RC4CryptoAPI", StringComparison.OrdinalIgnoreCase);

                            if (isDeprecated)
                            {
                                Console.WriteLine("WARNING: The workbook uses a deprecated encryption algorithm. " +
                                                  "Consider re‑encrypting it with a modern algorithm (e.g., AES256).");
                            }
                            else
                            {
                                Console.WriteLine("The encryption algorithm is up‑to‑date.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Unable to retrieve encryption information.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("EncryptionInfo property not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Failed to retrieve encryption details: {innerEx.Message}");
                }
            }
            catch (CellsException ex)
            {
                // Handles other Aspose.Cells related errors
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
