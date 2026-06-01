using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Directory containing the Excel workbooks
                string folderPath = @"C:\Workbooks";

                // Verify the directory exists
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                // Define the shared content type property name and value
                const string propertyName = "SharedProperty";
                const string propertyValue = "SharedValue";

                // Get all Excel files in the directory
                string[] files = Directory.GetFiles(folderPath, "*.xlsx");

                foreach (string filePath in files)
                {
                    // Ensure the file still exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook (will throw if the file is password‑protected)
                        Workbook workbook = new Workbook(filePath);

                        // Add the shared content type property
                        workbook.ContentTypeProperties.Add(propertyName, propertyValue);

                        // Overwrite the original file
                        workbook.Save(filePath);

                        Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        // Detect password‑protected files by message content
                        if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"Skipped password‑protected file: {Path.GetFileName(filePath)}");
                        }
                        else
                        {
                            Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Content type property added to all applicable workbooks.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors in the overall process
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}