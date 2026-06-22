using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionScanner
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory to scan – change as needed or pass as argument
            string folderPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder does not exist: {folderPath}");
                return;
            }

            // Get all .xlsx and .xls files (case‑insensitive)
            string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string file in files)
            {
                string ext = Path.GetExtension(file);
                if (!string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Skip non‑Excel files
                }

                try
                {
                    // Use Aspose.Cells utility to detect file format and encryption status
                    FileFormatInfo info = FileFormatUtil.DetectFileFormat(file);
                    bool encrypted = info.IsEncrypted;

                    Console.WriteLine($"{Path.GetFileName(file)} : Encrypted = {encrypted}");
                }
                catch (Exception ex)
                {
                    // If detection fails, report the error but continue processing other files
                    Console.WriteLine($"{Path.GetFileName(file)} : Error detecting encryption – {ex.Message}");
                }
            }
        }
    }
}