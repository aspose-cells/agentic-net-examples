// Title: Encrypt an Excel workbook as a multi‑page TIFF and upload it securely via FTPS (TLS) using C# and Aspose.Cells
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, saves it as a multi‑page TIFF in memory, encrypts the TIFF using AES‑256‑CBC with a password‑derived key, and uploads the encrypted bytes to an FTPS server using FtpWebRequest with TLS. | Show how to modify the encryption routine to use a custom salt and store the initialization vector in a separate file before performing the FTPS upload. | Create a streaming implementation that encrypts the workbook image on‑the‑fly and streams the encrypted data directly to the FTPS request without loading the entire TIFF into memory.
// Common Searches: C# convert Excel workbook to multi‑page TIFF with Aspose.Cells and upload via FTPS | How to encrypt a TIFF image using AES in .NET before sending over secure FTP | Upload encrypted files to an FTPS server using FtpWebRequest and TLS in C# | Save Excel as TIFF in memory and stream to FTPS without temporary files | Derive AES key from password for encrypting workbook images in C#
// Tags: Aspose.Cells export workbook to multi-page TIFF | AES-CBC image encryption in C# | FTPS file transfer with FtpWebRequest .NET | secure upload of encrypted Excel images | on‑the‑fly encryption and streaming to FTPS

using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using Aspose.Cells;

// The program loads an Excel workbook, renders it to a multi‑page TIFF in memory, encrypts the TIFF bytes with AES‑256‑CBC using a password‑derived key, and uploads the encrypted file to an FTPS server via TLS using FtpWebRequest.
class Program
{
    static void Main()
    {
        try
        {
            // Paths and credentials (replace with actual values)
            string workbookPath = @"C:\Input\Sample.xlsx";
            string tiffFileName = "WorkbookImage.tiff";
            string ftpServer = "ftps://secureftp.example.com/upload/";
            string ftpUser = "ftpUser";
            string ftpPassword = "ftpPassword";
            string encryptionPassword = "StrongEncryptionKey"; // password for AES encryption

            // Verify workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Render the workbook to a TIFF image in memory
            using (MemoryStream tiffStream = new MemoryStream())
            {
                // Save the entire workbook as a multi‑page TIFF
                workbook.Save(tiffStream, SaveFormat.Tiff);
                byte[] tiffBytes = tiffStream.ToArray();

                // Encrypt the TIFF bytes using AES
                byte[] encryptedBytes = EncryptData(tiffBytes, encryptionPassword);

                // Upload the encrypted file via FTPS (TLS)
                UploadFileToFtps(ftpServer + tiffFileName, encryptedBytes, ftpUser, ftpPassword);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Encrypt data with AES (CBC, PKCS7 padding)
    private static byte[] EncryptData(byte[] plainData, string password)
    {
        // Derive a 256‑bit key and a 128‑bit IV from the password
        using (var rfc2898 = new Rfc2898DeriveBytes(password, 16, 10000, HashAlgorithmName.SHA256))
        {
            byte[] key = rfc2898.GetBytes(32); // 256‑bit key
            byte[] iv = rfc2898.GetBytes(16);  // 128‑bit IV

            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (var encryptor = aes.CreateEncryptor())
                using (var ms = new MemoryStream())
                {
                    // Prepend IV for later decryption (optional)
                    ms.Write(iv, 0, iv.Length);
                    using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                    {
                        cs.Write(plainData, 0, plainData.Length);
                        cs.FlushFinalBlock();
                    }
                    return ms.ToArray();
                }
            }
        }
    }

    // Upload encrypted data to an FTPS server using TLS
    private static void UploadFileToFtps(string uri, byte[] data, string user, string password)
    {
        try
        {
            // Create the request
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(uri);
            request.Method = WebRequestMethods.Ftp.UploadFile;
            request.Credentials = new NetworkCredential(user, password);
            request.EnableSsl = true;               // Use TLS/SSL
            request.UseBinary = true;
            request.ContentLength = data.Length;

            // Write the encrypted data to the request stream
            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(data, 0, data.Length);
            }

            // Get the response to ensure the upload succeeded
            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                Console.WriteLine($"Upload status: {response.StatusDescription}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FTPS upload error: {ex.Message}");
        }
    }
}
