// Title: Check that PDF generated from Excel reflects updated linked shape after RefreshAll using Aspose.Cells for .NET
// Common Searches: how to confirm linked shape image is updated in PDF after workbook.RefreshAll in Aspose.Cells C# | compare PDF output before and after RefreshAll to verify linked shape changes with Aspose.Cells | C# compute SHA256 hash of PDF files to detect linked shape refresh in Aspose.Cells | validate external data connection refresh updates linked images in exported PDF using Aspose.Cells | Aspose.Cells verify PDF changes when linked shapes are refreshed in an Excel workbook
// Tags: Aspose.Cells RefreshAll PDF export verification | C# PDF hash comparison Aspose.Cells | linked shape image update detection Aspose.Cells | Excel to PDF conversion after data refresh Aspose.Cells | validate linked shape refresh in PDF output

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// The example loads an Excel workbook, saves a PDF before and after calling RefreshAll to update external data connections and linked shapes, computes SHA256 hashes for both PDFs, and compares the hashes to determine whether the linked shape image change is reflected in the exported PDF.
class VerifyLinkedShapePdfUpdate
{
    static void Main()
    {
        try
        {
            // Paths for the source Excel file and the PDF outputs
            string excelPath = @"C:\Temp\LinkedShapeWorkbook.xlsx";
            string pdfBeforeRefresh = @"C:\Temp\LinkedShapeBefore.pdf";
            string pdfAfterRefresh = @"C:\Temp\LinkedShapeAfter.pdf";

            // Ensure the source workbook exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Source Excel file not found: {excelPath}");
                return;
            }

            // 1. Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // 2. Save the initial PDF (before refreshing linked shapes)
            workbook.Save(pdfBeforeRefresh, SaveFormat.Pdf);

            // 3. Refresh all external data connections and linked shapes
            workbook.RefreshAll();

            // 4. Save the PDF after the refresh operation
            workbook.Save(pdfAfterRefresh, SaveFormat.Pdf);

            // 5. Verify that the PDF reflects the updated image by comparing file hashes
            string hashBefore = ComputeFileHash(pdfBeforeRefresh);
            string hashAfter = ComputeFileHash(pdfAfterRefresh);

            if (hashBefore != hashAfter)
                Console.WriteLine("PDF output has changed after refresh – updated linked shape image is reflected.");
            else
                Console.WriteLine("PDF output is unchanged – linked shape image may not have been refreshed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to compute SHA256 hash of a file
    private static string ComputeFileHash(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha = SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
    }
}
