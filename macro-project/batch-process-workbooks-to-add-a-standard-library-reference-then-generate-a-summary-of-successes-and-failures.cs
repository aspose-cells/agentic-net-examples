using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string sourceFolder = "InputWorkbooks";
        // Folder where processed workbooks will be saved
        string outputFolder = "OutputWorkbooks";

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all macro-enabled workbooks in the source folder
        string[] files = Directory.GetFiles(sourceFolder, "*.xlsm", SearchOption.TopDirectoryOnly);

        int successCount = 0;
        int failureCount = 0;

        foreach (string filePath in files)
        {
            try
            {
                // Load the workbook from file
                Workbook workbook = new Workbook(filePath);

                // Add the standard library reference if a VBA project is present
                if (workbook.VbaProject != null)
                {
                    workbook.VbaProject.References.AddRegisteredReference(
                        "stdole",
                        "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");
                }

                // Determine the output file path (overwrite in the output folder)
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the workbook preserving the macro-enabled format
                workbook.Save(outputPath, SaveFormat.Xlsm);

                successCount++;
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                failureCount++;
            }
        }

        // Output a summary of the batch operation
        Console.WriteLine($"Batch processing completed. Successes: {successCount}, Failures: {failureCount}");
    }
}