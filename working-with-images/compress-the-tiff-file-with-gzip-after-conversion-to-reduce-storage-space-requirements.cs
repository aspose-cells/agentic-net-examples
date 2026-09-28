// Title: Compress a multi‑page TIFF generated from an Excel workbook using GZip in C# with Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, saves the workbook as a multi‑page TIFF, and compresses the TIFF into a .gz file using GZipStream. | Show how to pipe the TIFF output directly into a GZipStream to avoid creating a temporary TIFF file on disk. | Create a reusable method that accepts a Workbook object and returns a byte array containing the GZip‑compressed TIFF data.
// Common Searches: asp.net convert excel workbook to multi page tiff and gzip the file | c# compress tiff produced by Aspose.Cells with GZipStream | how to delete temporary tiff after creating .tiff.gz archive in .NET
// Tags: Aspose.Cells export workbook to multi‑page TIFF | C# GZipStream compress TIFF image | remove temporary TIFF after GZip compression | store Excel as .tiff.gz archive | stream TIFF directly into GZip without intermediate file

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The example loads an Excel workbook, saves it as a multi‑page TIFF using Aspose.Cells, reads the TIFF bytes, compresses them into a GZip file with GZipStream, and optionally deletes the intermediate TIFF to reduce storage usage.
class TiffGzipCompression
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Save the workbook as a TIFF image
        // The SaveFormat.Tiff option creates a multi-page TIFF where each worksheet becomes a page
        workbook.Save("output.tiff", SaveFormat.Tiff);

        // Read the generated TIFF file into a byte array
        byte[] tiffBytes = File.ReadAllBytes("output.tiff");

        // Create a GZip compressed file
        using (FileStream compressedFile = new FileStream("output.tiff.gz", FileMode.Create))
        using (GZipStream gzipStream = new GZipStream(compressedFile, CompressionMode.Compress))
        {
            // Write the TIFF bytes into the GZip stream
            gzipStream.Write(tiffBytes, 0, tiffBytes.Length);
        }

        // Optional: delete the intermediate uncompressed TIFF if no longer needed
        // File.Delete("output.tiff");

        Console.WriteLine("TIFF file has been compressed to GZip successfully.");
    }
}
