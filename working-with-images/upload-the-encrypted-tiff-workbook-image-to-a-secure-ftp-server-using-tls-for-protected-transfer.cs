// Title: Encrypt Aspose.Cells Workbook, Render as LZW TIFF, and Upload via FTPS (TLS) in C#
// Description: Creates a Workbook, adds sample data, applies strong password protection and 128‑bit encryption, renders the first worksheet to a LZW‑compressed TIFF in a MemoryStream, and securely uploads the image to an FTPS server using TLS with FtpWebRequest.
// Keywords: Aspose.Cells | C# workbook encryption | LZW TIFF rendering | FTPS upload .NET | TLS secure FTP | FtpWebRequest | memory stream TIFF | Excel to image conversion | password protected workbook | secure file transfer
// Common Searches: How to encrypt an Excel workbook with Aspose.Cells C# | Render worksheet to LZW compressed TIFF using Aspose.Cells | Upload a file to FTPS server with TLS in C# | Stream TIFF directly to FTPS request without resetting position | Set 128‑bit encryption for Aspose.Cells workbook
// Developer Intent: Generate a password‑protected workbook, convert the first sheet to a compressed TIFF, and transfer the image securely over FTPS.
// Use Cases: Send confidential financial reports as encrypted TIFF images to partners via a TLS‑protected FTP site. | Archive sensitive Excel data by converting each workbook to an encrypted TIFF and storing it in a secure FTPS repository. | Integrate automated snapshot creation of workbooks into CI pipelines, rendering them as TIFFs and uploading to a protected FTPS server.
// AI Prompts: Add TLS 1.2 enforcement and custom certificate validation to the FTPS upload code. | Implement retry logic and detailed error handling for FtpWebResponse failures during the TIFF transfer. | Show how to pipe the rendered TIFF stream directly into the FTPS request without resetting the MemoryStream.

using System;
using System.IO;
using System.Net;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a Workbook, adds sample data, applies strong password protection and 128‑bit encryption, renders the first worksheet to a LZW‑compressed TIFF in a MemoryStream, and securely uploads the image to an FTPS server using TLS with FtpWebRequest.
class UploadEncryptedTiffViaFtps
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add sample data
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Secure FTPS Upload Demo");
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // 2. Encrypt the workbook with a password and strong encryption
                workbook.Settings.Password = "StrongPassword123";
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // 3. Render the first worksheet to a TIFF image in memory
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // TIFF-specific rendering options
                    TiffCompression = TiffCompression.CompressionLZW,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                SheetRender renderer = new SheetRender(sheet, imgOptions);
                using (MemoryStream tiffStream = new MemoryStream())
                {
                    // Render worksheet to TIFF
                    renderer.ToTiff(tiffStream);
                    tiffStream.Position = 0; // Reset stream for reading

                    // 4. Upload the TIFF stream to a secure FTP server using TLS (FTPS)
                    string ftpUrl = "ftp://your-secure-ftp-server.com/remote/path/encrypted_image.tiff";
                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpUrl);
                    request.Method = WebRequestMethods.Ftp.UploadFile;
                    request.EnableSsl = true; // Enforce TLS/SSL
                    request.Credentials = new NetworkCredential("ftpUser", "ftpPassword");
                    request.UseBinary = true;
                    request.KeepAlive = false;

                    // Write the TIFF data to the request stream
                    using (Stream requestStream = request.GetRequestStream())
                    {
                        tiffStream.CopyTo(requestStream);
                    }

                    // Get the response to ensure the upload succeeded
                    using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                    {
                        Console.WriteLine($"Upload status: {response.StatusDescription}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
