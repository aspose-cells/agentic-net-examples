// Title: Encrypt an Excel workbook to a TIFF image and transfer it securely via FTPS (TLS) using Aspose.Cells for .NET (C#)
// Common Searches: C# how to render an Aspose.Cells worksheet to a TIFF and upload it over FTPS | Encrypt a TIFF file with AES in C# before sending to a secure FTP server | Upload encrypted workbook image to FTPS using FtpWebRequest in .NET | Aspose.Cells generate LZW compressed TIFF and transfer via TLS FTP | Securely transfer encrypted Excel image to FTPS with Aspose.Cells and AES
// Tags: render worksheet to TIFF using Aspose.Cells | AES‑CBC encryption of TIFF in .NET | secure FTPS upload with FtpWebRequest | LZW compressed TIFF generation | transfer encrypted workbook image over TLS

using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The example creates a sample workbook, renders the first worksheet to a 24‑bit LZW‑compressed TIFF at 300 dpi, encrypts the TIFF using a 256‑bit AES key in CBC mode, and uploads the encrypted file to an FTPS server via FtpWebRequest with SSL/TLS enabled.
class Program
{
    // Configuration – replace with your actual values
    private const string FtpHost = "ftps://your-secure-ftp-server.com";
    private const string FtpUsername = "ftp_user";
    private const string FtpPassword = "ftp_password";
    private const string RemoteFileName = "encrypted_workbook.tiff";

    // AES encryption key and IV – in a real scenario store securely
    private static readonly byte[] AesKey = new byte[32] // 256‑bit key
    {
        0x10,0x20,0x30,0x40,0x50,0x60,0x70,0x80,
        0x90,0xA0,0xB0,0xC0,0xD0,0xE0,0xF0,0x01,
        0x12,0x23,0x34,0x45,0x56,0x67,0x78,0x89,
        0x9A,0xAB,0xBC,0xCD,0xDE,0xEF,0xF1,0x02
    };
    private static readonly byte[] AesIv = new byte[16] // 128‑bit IV
    {
        0x01,0x02,0x03,0x04,0x05,0x06,0x07,0x08,
        0x09,0x0A,0x0B,0x0C,0x0D,0x0E,0x0F,0x10
    };

    static void Main()
    {
        try
        {
            // 1. Create a sample workbook
            Workbook workbook = CreateSampleWorkbook();

            // 2. Render the workbook (first sheet) to a TIFF image file
            string tiffPath = Path.Combine(Path.GetTempPath(), "workbook.tiff");
            RenderWorkbookToTiff(workbook, tiffPath);

            // 3. Encrypt the TIFF file using AES
            string encryptedPath = Path.Combine(Path.GetTempPath(), "workbook_encrypted.tiff");
            EncryptFile(tiffPath, encryptedPath);

            // 4. Upload the encrypted file via FTPS (TLS)
            UploadFileToFtps(encryptedPath, RemoteFileName);

            Console.WriteLine("Upload completed successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    private static Workbook CreateSampleWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        ws.Cells["A1"].PutValue("Sample Data");
        ws.Cells["A2"].PutValue(DateTime.Now);
        ws.Cells["B1"].PutValue(12345);
        ws.Cells["B2"].PutValue(67.89);

        // Add a few more rows to make the image a bit larger
        for (int i = 3; i <= 20; i++)
        {
            ws.Cells[$"A{i}"].PutValue($"Row {i}");
            ws.Cells[$"B{i}"].PutValue(i * 10);
        }

        return wb;
    }

    private static void RenderWorkbookToTiff(Workbook workbook, string outputPath)
    {
        // Configure image options for TIFF
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            TiffCompression = TiffCompression.CompressionLZW,
            TiffColorDepth = ColorDepth.Format24bpp,
            HorizontalResolution = 300,
            VerticalResolution = 300,
            OnePagePerSheet = true // render each sheet as a single page
        };

        // Render the first worksheet (index 0) to TIFF
        Worksheet sheet = workbook.Worksheets[0];
        SheetRender renderer = new SheetRender(sheet, options);
        // ToTiff renders the whole worksheet to a single TIFF file
        renderer.ToTiff(outputPath);
        Console.WriteLine($"TIFF image generated at: {outputPath}");
    }

    private static void EncryptFile(string inputFile, string outputFile)
    {
        using (FileStream fsInput = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
        using (FileStream fsEncrypted = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
        using (Aes aes = Aes.Create())
        {
            aes.Key = AesKey;
            aes.IV = AesIv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;

            using (CryptoStream cryptoStream = new CryptoStream(fsEncrypted, aes.CreateEncryptor(), CryptoStreamMode.Write))
            {
                fsInput.CopyTo(cryptoStream);
            }
        }

        Console.WriteLine($"File encrypted to: {outputFile}");
    }

    private static void UploadFileToFtps(string localFilePath, string remoteFileName)
    {
        // Build the full URI (ensure it ends with the file name)
        string uri = $"{FtpHost.TrimEnd('/')}/{remoteFileName}";

        // Create the request
        FtpWebRequest request = (FtpWebRequest)WebRequest.Create(uri);
        request.Method = WebRequestMethods.Ftp.UploadFile;
        request.Credentials = new NetworkCredential(FtpUsername, FtpPassword);
        request.EnableSsl = true; // TLS/SSL
        request.UseBinary = true;
        request.KeepAlive = false;

        // Read the file data
        byte[] fileContents;
        using (FileStream sourceStream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read))
        {
            fileContents = new byte[sourceStream.Length];
            sourceStream.Read(fileContents, 0, fileContents.Length);
        }

        // Upload the data
        request.ContentLength = fileContents.Length;
        using (Stream requestStream = request.GetRequestStream())
        {
            requestStream.Write(fileContents, 0, fileContents.Length);
        }

        // Get the response to ensure the upload succeeded
        using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
        {
            Console.WriteLine($"Upload status: {response.StatusDescription}");
        }
    }
}
