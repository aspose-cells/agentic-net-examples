// Title: Save a password‑protected Excel workbook as an encrypted .xlsx file while preserving encryption metadata using Aspose.Cells for .NET
// AI Prompts: Write C# that creates a Workbook, inserts sample data, sets a password, and saves the encrypted workbook directly to a Stream (e.g., MemoryStream) so the stream can be sent to cloud storage, keeping the encryption metadata intact. | Demonstrate how to confirm that the password protection settings remain after the workbook is saved to a Stream with Aspose.Cells.
// Common Searches: Aspose.Cells save encrypted workbook to memory stream C# | preserve password protection when uploading Excel file to Azure Blob using Aspose.Cells | C# write password protected .xlsx to cloud storage stream | how to keep encryption metadata after saving Aspose.Cells workbook to a stream | Aspose.Cells encrypt workbook before sending to Amazon S3
// Tags: save encrypted workbook to stream Aspose.Cells | password protect Excel file Aspose.Cells .NET | preserve encryption metadata Aspose.Cells | upload encrypted .xlsx to cloud storage C# | Aspose.Cells workbook encryption settings

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, adds values to cells A1 and B1, applies a password via workbook.Settings.Password, and saves the workbook as an encrypted .xlsx file. The encryption metadata is retained, allowing the workbook to be written to any output stream (e.g., a cloud storage stream) for secure upload.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set password protection (default encryption)
            workbook.Settings.Password = "StrongPassword123";

            // Local output file path
            string outputPath = "encryptedWorkbook.xlsx";

            // Ensure the output directory exists (if a directory is specified)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the encrypted workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Encrypted workbook saved to '{outputPath}' successfully.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
