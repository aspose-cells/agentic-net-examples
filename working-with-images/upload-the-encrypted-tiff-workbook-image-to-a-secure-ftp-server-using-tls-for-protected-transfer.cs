// Title: Convert an Excel workbook to an encrypted TIFF and upload it securely via FTPS using C# and Aspose.Cells
// AI Prompts: Generate C# code that loads a .xlsx file, saves it as a TIFF with Aspose.Cells, encrypts the TIFF using AES‑256‑CBC, and uploads the encrypted file to an FTPS server with TLS. | Show how to derive an AES key from a password with Rfc2898DeriveBytes and apply it to encrypt a TIFF before sending it over FTPS in C#. | Provide a C# snippet that deletes the temporary TIFF files after a successful FTPS upload and logs the FTP response.
// Common Searches: c# export excel to tiff then encrypt and upload via ftps | how to use aspocells to save workbook as tiff in .net core | ftps upload encrypted image file using FtpWebRequest c# | aes 256 cbc encryption of tiff file in c# example | delete temporary files after ftps upload c#
// Tags: aspocells save workbook as tiff | aes cbc encryption of tiff c# | secure ftps transfer of encrypted tiff via FtpWebRequest | remove temporary tiff files post ftps upload | derive aes key from password c#

using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using Aspose.Cells;

// // This program creates a sample Excel workbook if missing, saves it as a TIFF using Aspose.Cells, encrypts the TIFF with AES‑256‑CBC, uploads the encrypted file to an FTPS server via FtpWebRequest, and then deletes the temporary files.
class Program
{
    static void Main()
    {
        // Paths for the source Excel file, intermediate TIFF, and encrypted TIFF
        string excelPath = @"C:\Data\Workbook.xlsx";
        string tiffPath = @"C:\Data\Workbook.tiff";
        string encryptedTiffPath = @"C:\Data\Workbook_encrypted.tiff";

        // FTP server details (replace with actual values)
        string ftpHost = "ftps.example.com";
        string ftpUser = "ftpUser";
        string ftpPassword = "ftpPassword";
        string remoteFilePath = "/uploads/Workbook_encrypted.tiff";

        try
        {
            // -----------------------------------------------------------------
            // 1. Ensure the Excel workbook exists; create a simple one if missing
            // -----------------------------------------------------------------
            if (!File.Exists(excelPath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(excelPath));
                Workbook newWb = new Workbook();
                newWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                newWb.Save(excelPath);
                Console.WriteLine($"Created placeholder workbook at '{excelPath}'.");
            }

            // Load the workbook and export it as a TIFF image
            Workbook workbook = new Workbook(excelPath);
            workbook.Save(tiffPath, SaveFormat.Tiff);
            Console.WriteLine($"Workbook exported to TIFF at '{tiffPath}'.");

            // -----------------------------------------------------------------
            // 2. Encrypt the TIFF file using AES (CBC mode with PKCS7 padding)
            // -----------------------------------------------------------------
            byte[] key = new byte[32]; // 256‑bit key
            byte[] iv = new byte[16];  // 128‑bit IV
            for (int i = 0; i < key.Length; i++) key[i] = (byte)i;
            for (int i = 0; i < iv.Length; i++) iv[i] = (byte)(i + 1);

            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (FileStream inputFile = new FileStream(tiffPath, FileMode.Open, FileAccess.Read))
                using (FileStream outputFile = new FileStream(encryptedTiffPath, FileMode.Create, FileAccess.Write))
                using (CryptoStream cryptoStream = new CryptoStream(outputFile, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    inputFile.CopyTo(cryptoStream);
                }
            }
            Console.WriteLine($"Encrypted TIFF saved to '{encryptedTiffPath}'.");

            // -----------------------------------------------------------------
            // 3. Upload the encrypted TIFF to a secure FTP server using TLS (FTPS)
            // -----------------------------------------------------------------
            string uri = $"ftps://{ftpHost}{remoteFilePath}";
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(uri);
            request.Method = WebRequestMethods.Ftp.UploadFile;
            request.Credentials = new NetworkCredential(ftpUser, ftpPassword);
            request.EnableSsl = true;               // Enforce TLS
            request.UseBinary = true;
            request.KeepAlive = false;

            byte[] fileContents = File.ReadAllBytes(encryptedTiffPath);
            request.ContentLength = fileContents.Length;

            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(fileContents, 0, fileContents.Length);
            }

            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                Console.WriteLine($"Upload status: {response.StatusDescription}");
            }

            // Cleanup temporary files (optional)
            File.Delete(tiffPath);
            File.Delete(encryptedTiffPath);
            Console.WriteLine("Temporary files deleted.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
