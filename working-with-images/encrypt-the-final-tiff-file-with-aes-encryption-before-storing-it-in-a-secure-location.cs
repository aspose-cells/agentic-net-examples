// Title: Encrypt a TIFF image generated from an Excel worksheet with AES‑256 using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render worksheet to TIFF and encrypt with AES-256 | How to apply password‑based AES encryption to a TIFF file in .NET | Save Excel sheet as TIFF then secure it with AES CBC in C# | Encrypt generated TIFF image using PBKDF2 derived key in Aspose.Cells example | Store encrypted TIFF in a protected folder using Aspose.Cells and System.Security.Cryptography
// Tags: Aspose.Cells render worksheet to TIFF | AES-256 CBC encryption of TIFF in C# | PBKDF2 password derived key for file encryption | CryptoStream file encryption .NET | Secure folder storage for encrypted files

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffAesEncryption
{
    // The sample loads an Excel workbook, renders the first worksheet to an LZW‑compressed 24‑bit TIFF with Aspose.Cells, then encrypts the TIFF using AES‑256 CBC (key and IV derived from a password via PBKDF2), writes the salt and ciphertext to a secure directory, and finally removes the plaintext TIFF.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed or pass via command‑line arguments
            string inputExcelPath = "input.xlsx";
            string tiffPath = "output.tiff";
            string encryptedTiffPath = "output.tiff.enc";
            string secureFolder = "SecureFolder";

            // Password used for AES encryption – in a real scenario obtain securely (e.g., from a vault)
            string encryptionPassword = "StrongPassword!123";

            try
            {
                // Ensure the secure folder exists
                if (!Directory.Exists(secureFolder))
                {
                    Directory.CreateDirectory(secureFolder);
                }

                // -----------------------------------------------------------------
                // 1. Load the workbook and render the first worksheet to a TIFF file
                // -----------------------------------------------------------------
                if (!File.Exists(inputExcelPath))
                {
                    Console.WriteLine($"Error: Input Excel file not found at '{inputExcelPath}'.");
                    return;
                }

                Workbook workbook = new Workbook(inputExcelPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Configure TIFF rendering options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,
                    TiffCompression = TiffCompression.CompressionLZW,
                    TiffColorDepth = ColorDepth.Format24bpp,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Render the worksheet to TIFF
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                renderer.ToTiff(tiffPath);
                Console.WriteLine($"TIFF file generated at '{tiffPath}'.");

                // -----------------------------------------------------------------
                // 2. Encrypt the generated TIFF using AES (CBC, 256‑bit key)
                // -----------------------------------------------------------------
                string encryptedFileFullPath = Path.Combine(secureFolder, encryptedTiffPath);
                EncryptFileAes(tiffPath, encryptedFileFullPath, encryptionPassword);
                Console.WriteLine($"Encrypted TIFF saved to secure location: '{encryptedFileFullPath}'.");

                // Optionally delete the plain TIFF after successful encryption
                try
                {
                    File.Delete(tiffPath);
                    Console.WriteLine("Plain TIFF file deleted after encryption.");
                }
                catch (Exception exDel)
                {
                    Console.WriteLine($"Warning: Could not delete plain TIFF file. {exDel.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="inputFilePath">Path to the plaintext file.</param>
        /// <param name="outputFilePath">Path where the encrypted file will be written.</param>
        /// <param name="password">Password used to derive the encryption key.</param>
        private static void EncryptFileAes(string inputFilePath, string outputFilePath, string password)
        {
            // Parameters for key derivation
            const int saltSize = 16; // 128‑bit salt
            const int keySize = 32;  // 256‑bit key
            const int ivSize = 16;   // 128‑bit IV
            const int iterations = 100_000;

            // Generate a random salt
            byte[] salt = new byte[saltSize];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(salt);
            }

            // Derive key and IV from password + salt
            using (var keyDerivation = new Rfc2898DeriveBytes(password, salt, iterations, HashAlgorithmName.SHA256))
            {
                byte[] key = keyDerivation.GetBytes(keySize);
                byte[] iv = keyDerivation.GetBytes(ivSize);

                using (Aes aes = Aes.Create())
                {
                    aes.KeySize = keySize * 8;
                    aes.BlockSize = ivSize * 8;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.Key = key;
                    aes.IV = iv;

                    using (FileStream inputFs = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
                    using (FileStream outputFs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                    {
                        // Write the salt first – needed for decryption
                        outputFs.Write(salt, 0, salt.Length);

                        using (CryptoStream cryptoStream = new CryptoStream(outputFs, aes.CreateEncryptor(), CryptoStreamMode.Write))
                        {
                            inputFs.CopyTo(cryptoStream);
                        }
                    }
                }
            }
        }
    }
}
