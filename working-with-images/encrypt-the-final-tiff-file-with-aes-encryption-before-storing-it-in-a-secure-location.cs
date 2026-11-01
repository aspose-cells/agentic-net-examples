// Title: Encrypt a TIFF image generated from an Aspose.Cells workbook with AES‑256‑CBC in C# and store it in a protected folder
// Common Searches: C# how to encrypt a TIFF file produced by Aspose.Cells using AES‑256‑CBC | save an AES encrypted TIFF image from Aspose.Cells to a secure directory in .NET | example of encrypting TIFF bytes with System.Security.Cryptography in C# | store IV and key alongside encrypted TIFF data using Aes class | Aspose.Cells export workbook to TIFF then apply AES encryption before saving
// Tags: AES-256-CBC encryption of TIFF in C# | Aspose.Cells workbook to TIFF conversion | Encrypt TIFF bytes with System.Security.Cryptography | Write IV and key to encrypted file in .NET | Secure storage of encrypted image files

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// The sample creates an Aspose.Cells workbook, saves it as a TIFF image to a memory stream, encrypts the TIFF bytes with AES‑256‑CBC (generating a random key and IV), and writes the IV, key, and ciphertext to a file in a protected directory.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook (using Aspose.Cells lifecycle rule)
            Workbook workbook = new Workbook();

            // Populate the workbook with sample data (replace with actual data as needed)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // 2. Save the workbook as a TIFF image into a memory stream
            using (MemoryStream tiffStream = new MemoryStream())
            {
                workbook.Save(tiffStream, SaveFormat.Tiff);
                byte[] tiffBytes = tiffStream.ToArray();

                // 3. Encrypt the TIFF bytes using AES
                byte[] encryptedBytes = EncryptWithAes(tiffBytes, out byte[] iv, out byte[] key);

                // 4. Store the encrypted file in a secure location
                string outputPath = @"C:\SecureLocation\encrypted_output.tiff";

                // Ensure the target directory exists
                string directory = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write IV, Key and encrypted data to the file
                using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(iv, 0, iv.Length);   // IV (16 bytes)
                    fs.Write(key, 0, key.Length); // Key (32 bytes for AES‑256)
                    fs.Write(encryptedBytes, 0, encryptedBytes.Length);
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Encrypts data using AES‑256‑CBC. Returns ciphertext and outputs the IV and Key used.
    private static byte[] EncryptWithAes(byte[] plainData, out byte[] iv, out byte[] key)
    {
        using (Aes aes = Aes.Create())
        {
            aes.KeySize = 256;               // AES‑256
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;

            // Generate a random key and IV
            aes.GenerateKey();
            aes.GenerateIV();

            key = aes.Key;
            iv = aes.IV;

            using (ICryptoTransform encryptor = aes.CreateEncryptor())
            using (MemoryStream msCipher = new MemoryStream())
            using (CryptoStream cs = new CryptoStream(msCipher, encryptor, CryptoStreamMode.Write))
            {
                cs.Write(plainData, 0, plainData.Length);
                cs.FlushFinalBlock();
                return msCipher.ToArray();
            }
        }
    }
}
