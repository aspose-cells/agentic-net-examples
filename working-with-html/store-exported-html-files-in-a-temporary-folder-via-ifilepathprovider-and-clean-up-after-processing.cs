using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExportTemp
{
    // Provides a full path for each worksheet HTML file inside a temporary folder.
    internal class TempFilePathProvider : IFilePathProvider
    {
        private readonly string _tempFolder;

        public TempFilePathProvider(string tempFolder)
        {
            _tempFolder = tempFolder;
        }

        public string GetFullName(string sheetName)
        {
            // Ensure the temporary folder exists.
            Directory.CreateDirectory(_tempFolder);
            // Example: C:\Temp\AsposeHtmlExport\Sheet1.html
            return Path.Combine(_tempFolder, $"{sheetName}.html");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a unique temporary directory for this export operation.
            string tempExportFolder = Path.Combine(Path.GetTempPath(),
                "AsposeHtmlExport_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempExportFolder);

            try
            {
                // Create a workbook and add sample data.
                Workbook workbook = new Workbook();

                // First worksheet (index 0 already exists).
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["A2"].PutValue("World");

                // Add a second worksheet and obtain its reference.
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[newSheetIndex];
                sheet2.Name = "Sheet2";
                sheet2.Cells["B1"].PutValue("Another sheet");

                // Configure HtmlSaveOptions to use the custom temporary file path provider.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    FilePathProvider = new TempFilePathProvider(tempExportFolder),
                    IsExpImageToTempDir = true // Export images to a temporary directory.
                };

                // Save the main HTML file (references the per‑sheet files generated above).
                string mainHtmlPath = Path.Combine(tempExportFolder, "Workbook.html");
                workbook.Save(mainHtmlPath, saveOptions);

                Console.WriteLine($"Main HTML saved to: {mainHtmlPath}");
                Console.WriteLine($"Per‑sheet HTML files are stored in: {tempExportFolder}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary folder and all its contents after processing.
                if (Directory.Exists(tempExportFolder))
                {
                    try
                    {
                        Directory.Delete(tempExportFolder, true);
                        Console.WriteLine($"Temporary folder '{tempExportFolder}' deleted.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to delete temporary folder: {ex.Message}");
                    }
                }
            }
        }
    }
}