// Title: Verify worksheet PNG integrity by computing and comparing its SHA‑256 hash with Aspose.Cells in C#
// Common Searches: C# compute SHA‑256 checksum of a PNG rendered from an Aspose.Cells worksheet | How to validate the integrity of a worksheet PNG image using Aspose.Cells .NET | Command‑line tool to compare expected SHA‑256 hash with generated PNG from Aspose.Cells | Generate a sample PNG from a workbook and verify its checksum in C# | Aspose.Cells render worksheet to PNG and check file hash for tamper detection
// Tags: SHA-256 verification of Aspose.Cells generated PNG | Aspose.Cells worksheet rendering to PNG using ImageOrPrintOptions | C# file hash calculation with System.Security.Cryptography | checksum comparison for rendered worksheet image | sample PNG generation fallback for missing file

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace VerifyPngChecksum
{
    // This C# console example uses Aspose.Cells to render the first worksheet of a workbook to a PNG file, computes the file’s SHA‑256 hash with System.Security.Cryptography, and optionally compares it to an expected checksum supplied via command‑line arguments, generating a sample PNG when the target file is missing.
    class Program
    {
        static int Main(string[] args)
        {
            try
            {
                string pngPath;
                string expectedChecksum = null;

                // Determine operation mode based on supplied arguments.
                if (args.Length == 0)
                {
                    // No arguments: generate a sample PNG, compute its checksum and display it.
                    pngPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
                    Console.WriteLine("No arguments supplied. Generating sample PNG at:");
                    Console.WriteLine($"  {pngPath}");
                    GenerateSamplePng(pngPath);
                }
                else if (args.Length == 1)
                {
                    // One argument: treat it as the PNG path, generate if missing, then just compute checksum.
                    pngPath = args[0];
                    Console.WriteLine("Only PNG path supplied. Expected checksum not provided; will only compute and display checksum.");
                }
                else if (args.Length == 2)
                {
                    // Two arguments: normal verification mode.
                    pngPath = args[0];
                    string expectedInput = args[1];

                    // Resolve expected checksum: read from file if it exists, otherwise treat as literal.
                    if (File.Exists(expectedInput))
                    {
                        expectedChecksum = File.ReadAllText(expectedInput).Trim();
                        Console.WriteLine($"Read expected checksum from file \"{expectedInput}\": {expectedChecksum}");
                    }
                    else
                    {
                        expectedChecksum = expectedInput.Trim();
                        Console.WriteLine($"Using expected checksum supplied directly: {expectedChecksum}");
                    }
                }
                else
                {
                    Console.WriteLine("Error: Too many arguments.");
                    Console.WriteLine("Usage: VerifyPngChecksum.exe <pngFilePath> [expectedChecksumOrFile]");
                    return 1;
                }

                // Ensure the PNG file exists; generate a sample if it does not.
                if (!File.Exists(pngPath))
                {
                    Console.WriteLine($"PNG file \"{pngPath}\" does not exist. Generating a sample PNG...");
                    try
                    {
                        string directory = Path.GetDirectoryName(pngPath);
                        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }

                        GenerateSamplePng(pngPath);
                        Console.WriteLine($"Sample PNG generated at \"{pngPath}\".");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to generate sample PNG: {ex.Message}");
                        return 99;
                    }
                }

                // Compute SHA‑256 checksum of the PNG file.
                string actualChecksum;
                try
                {
                    actualChecksum = ComputeSha256(pngPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error computing checksum: {ex.Message}");
                    return 99;
                }

                Console.WriteLine($"Computed SHA‑256 checksum: {actualChecksum}");

                // If an expected checksum was supplied, compare them.
                if (!string.IsNullOrEmpty(expectedChecksum))
                {
                    if (string.Equals(actualChecksum, expectedChecksum, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("Success: The PNG file integrity is verified. Checksums match.");
                        return 0;
                    }
                    else
                    {
                        Console.WriteLine("Failure: Checksum mismatch!");
                        Console.WriteLine($"Expected: {expectedChecksum}");
                        Console.WriteLine($"Actual  : {actualChecksum}");
                        return 2;
                    }
                }

                // No expected checksum to compare; just exit successfully after displaying the computed value.
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return 99;
            }
        }

        // Generates a simple PNG image from a workbook worksheet.
        private static void GenerateSamplePng(string outputPath)
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data.
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Set page setup to fit the data on a single page.
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.PrintArea = "A1:E10";
            pageSetup.SetFitToPages(1, 1);

            // Configure rendering options for PNG output.
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png,
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the first page to the specified PNG file.
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, outputPath);
        }

        // Computes the SHA‑256 hash of a file and returns it as a hex string.
        private static string ComputeSha256(string filePath)
        {
            using (FileStream stream = File.OpenRead(filePath))
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(stream);
                char[] chars = new char[hashBytes.Length * 2];
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    int b = hashBytes[i] >> 4;
                    chars[i * 2] = (char)(b > 9 ? b + 87 : b + 48);
                    b = hashBytes[i] & 0xF;
                    chars[i * 2 + 1] = (char)(b > 9 ? b + 87 : b + 48);
                }
                return new string(chars);
            }
        }
    }
}
