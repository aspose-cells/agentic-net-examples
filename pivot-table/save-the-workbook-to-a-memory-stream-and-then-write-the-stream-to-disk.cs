using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SaveWorkbookToMemoryStreamAndDisk
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data
                using (Workbook workbook = new Workbook())
                {
                    Worksheet worksheet = workbook.Worksheets[0];
                    worksheet.Cells["A1"].PutValue("Hello");
                    worksheet.Cells["B1"].PutValue("World");

                    // Save the workbook to a memory stream in XLSX format
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        workbook.Save(memoryStream, SaveFormat.Xlsx);

                        // Reset the stream position before reading
                        memoryStream.Position = 0;

                        string outputPath = "WorkbookFromMemoryStream.xlsx";

                        // Ensure we don't attempt to overwrite a locked file
                        if (File.Exists(outputPath))
                        {
                            File.Delete(outputPath);
                        }

                        // Write the memory stream content to a physical file
                        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                        {
                            memoryStream.CopyTo(fileStream);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public static class Program
    {
        public static void Main(string[] args)
        {
            SaveWorkbookToMemoryStreamAndDisk.Run();
        }
    }
}