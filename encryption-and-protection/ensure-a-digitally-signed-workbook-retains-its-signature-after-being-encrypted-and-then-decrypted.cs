// Title: How to keep a digital signature intact when encrypting and decrypting an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# encrypt an Excel workbook with a password using Aspose.Cells and preserve its digital signature | Aspose.Cells verify digital signature after saving an encrypted XLSX file | load password‑protected XLSX in Aspose.Cells and check digital signature count | use reflection to set OoxmlSaveOptions encryption properties in Aspose.Cells | preserve digital signature when re‑saving a protected workbook with Aspose.Cells
// Tags: Aspose.Cells encrypt workbook with password preserving digital signature | C# OoxmlSaveOptions strong encryption Aspose.Cells | LoadOptions password protected XLSX Aspose.Cells | digital signature collection count Aspose.Cells | reflection set encryption properties Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads a signed Excel workbook, encrypts it by applying a password and strong AES encryption via OoxmlSaveOptions (using reflection when needed), saves the encrypted file, then reloads it with LoadOptions using the same password, saves a decrypted copy, and finally compares the digital signature count before and after the round‑trip to confirm that the signature remains intact.
class DigitalSignaturePreservation
{
    static void Main()
    {
        try
        {
            // Path to the original workbook that already contains a digital signature
            const string originalPath = "SignedWorkbook.xlsx";

            // Ensure the source file exists
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Error: File \"{originalPath}\" not found.");
                return;
            }

            // Load the signed workbook
            Workbook signedWorkbook = new Workbook(originalPath);

            // Try to obtain the digital signature count via reflection (API may be absent)
            int originalSignatureCount = GetDigitalSignatureCount(signedWorkbook);
            Console.WriteLine($"Original signature count: {originalSignatureCount}");

            // --------------------------------------------------------------------
            // Encrypt the workbook by saving it with a password (AES 256 encryption)
            // --------------------------------------------------------------------
            const string password = "StrongPassword123";

            // Create save options for XLSX
            OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Attempt to set encryption properties via reflection
            bool encryptionSupported = SetPropertyIfExists(encryptOptions, "EncryptionType", Enum.Parse(typeof(EncryptionType), "StrongEncryption"));
            bool passwordSupported = SetPropertyIfExists(encryptOptions, "Password", password);

            if (!encryptionSupported || !passwordSupported)
            {
                Console.WriteLine("Encryption options are not supported in this Aspose.Cells version; saving without encryption.");
            }

            // Save the (potentially) encrypted workbook
            const string encryptedPath = "EncryptedWorkbook.xlsx";
            signedWorkbook.Save(encryptedPath, encryptOptions);
            Console.WriteLine($"Workbook saved to: {encryptedPath}");

            // --------------------------------------------------------------------
            // Decrypt the workbook by loading it with the password and re‑saving
            // --------------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();

            // Attempt to set the password for loading via reflection
            bool loadPasswordSupported = SetPropertyIfExists(loadOptions, "Password", password);
            if (!loadPasswordSupported)
            {
                Console.WriteLine("LoadOptions.Password is not supported; attempting to load without explicit password.");
            }

            // Load the encrypted workbook
            Workbook encryptedWorkbook = new Workbook(encryptedPath, loadOptions);

            // Save the decrypted workbook (no password)
            const string decryptedPath = "DecryptedWorkbook.xlsx";
            encryptedWorkbook.Save(decryptedPath);
            Console.WriteLine($"Workbook decrypted and saved to: {decryptedPath}");

            // --------------------------------------------------------------------
            // Verify that the digital signature is still present after the round‑trip
            // --------------------------------------------------------------------
            if (!File.Exists(decryptedPath))
            {
                Console.WriteLine($"Error: Decrypted file \"{decryptedPath}\" not found.");
                return;
            }

            Workbook finalWorkbook = new Workbook(decryptedPath);
            int finalSignatureCount = GetDigitalSignatureCount(finalWorkbook);
            Console.WriteLine($"Final signature count after encryption/decryption: {finalSignatureCount}");

            // Simple check – the counts should match when signature data is available
            if (originalSignatureCount >= 0 && finalSignatureCount >= 0)
            {
                if (originalSignatureCount == finalSignatureCount && originalSignatureCount > 0)
                {
                    Console.WriteLine("Success: Digital signature retained after encryption and decryption.");
                }
                else
                {
                    Console.WriteLine("Failure: Digital signature was lost during the process.");
                }
            }
            else
            {
                Console.WriteLine("Signature verification could not be performed due to missing API support.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Attempts to set a property via reflection; returns true if successful.
    private static bool SetPropertyIfExists(object target, string propertyName, object value)
    {
        try
        {
            PropertyInfo prop = target.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(target, value);
                return true;
            }
        }
        catch
        {
            // Ignored – property not available or cannot be set.
        }
        return false;
    }

    // Retrieves the digital signature count using reflection; returns -1 if unavailable.
    private static int GetDigitalSignatureCount(Workbook workbook)
    {
        try
        {
            PropertyInfo sigProp = workbook.GetType().GetProperty("DigitalSignatureCollection", BindingFlags.Public | BindingFlags.Instance);
            if (sigProp != null)
            {
                object collection = sigProp.GetValue(workbook);
                if (collection != null)
                {
                    PropertyInfo countProp = collection.GetType().GetProperty("Count", BindingFlags.Public | BindingFlags.Instance);
                    if (countProp != null)
                    {
                        object countObj = countProp.GetValue(collection);
                        if (countObj is int count)
                            return count;
                    }
                }
            }
        }
        catch
        {
            // Ignored – API not present.
        }
        return -1; // Indicates that signature information could not be retrieved.
    }
}
