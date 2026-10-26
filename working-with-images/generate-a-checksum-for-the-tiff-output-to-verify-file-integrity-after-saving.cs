// Title: Generate a SHA256 checksum for a TIFF image saved with Aspose.Cells in C#
// Common Searches: C# compute SHA256 hash of a TIFF file created by Aspose.Cells | verify TIFF image integrity after Aspose.Cells workbook.Save in .NET | how to calculate file checksum for an Aspose.Cells exported TIFF using SHA256 | Aspose.Cells export worksheet to TIFF and validate file hash in C# | generate file hash for image output from Aspose.Cells in .NET
// Tags: compute SHA256 checksum for TIFF using .NET | Aspose.Cells export worksheet to TIFF | verify TIFF file integrity in C# | C# file hash generation for image files | Aspose.Cells workbook.Save checksum validation | SHA256 hash of exported image

using System;
using Aspose.Cells;
using System.IO;
using System.Security.Cryptography;
using System.Text;

// Creates a workbook, populates data, saves it as a TIFF image with Aspose.Cells, then computes and prints a SHA256 checksum of the saved TIFF file to confirm its integrity.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(85);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(92);

        // Define the output TIFF file path
        string tiffPath = "output.tiff";

        // Save the workbook as a TIFF image
        workbook.Save(tiffPath, SaveFormat.Tiff);

        // Generate a SHA256 checksum for the saved TIFF file
        string checksum = ComputeSHA256Checksum(tiffPath);

        // Output the checksum to verify file integrity
        Console.WriteLine($"SHA256 checksum of '{tiffPath}': {checksum}");
    }

    // Helper method to compute SHA256 checksum of a file
    static string ComputeSHA256Checksum(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha = SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(stream);
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hash)
                sb.Append(b.ToString("x2"));
            return sb.ToString();
        }
    }
}
