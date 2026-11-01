// Title: C# batch job to process up to 100 Excel workbooks with Aspose.Cells: add optional custom metadata and remove unused styles
// Common Searches: how to batch process up to 100 Excel workbooks using Aspose.Cells in C# | add a custom document property to multiple Excel files with Aspose.Cells C# | remove unused styles from a workbook programmatically with Aspose.Cells | save processed Excel workbooks to a separate folder using Aspose.Cells C# | limit Aspose.Cells workbook processing loop to first 100 files in C#
// Tags: excel batch processing Aspose.Cells C# | custom document property addition Aspose.Cells | unused style cleanup Aspose.Cells | process first 100 workbooks limit | output directory saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchJob
{
    // The WorkbookProcessor class loads up to 100 Excel workbooks from a source folder, optionally adds a custom document property, removes any unused styles, and saves each workbook to a specified output directory using Aspose.Cells in C#.
    public class WorkbookProcessor
    {
        // Path to the folder containing source workbooks
        private readonly string _sourceFolder;
        // Path to the folder where processed workbooks will be saved
        private readonly string _outputFolder;
        // Optional metadata to add (key/value). Set to null if not needed.
        private readonly (string Key, string Value)? _optionalMetadata;

        public WorkbookProcessor(string sourceFolder, string outputFolder, (string Key, string Value)? optionalMetadata = null)
        {
            _sourceFolder = sourceFolder;
            _outputFolder = outputFolder;
            _optionalMetadata = optionalMetadata;
        }

        public void ProcessAll()
        {
            // Ensure output directory exists
            Directory.CreateDirectory(_outputFolder);

            // Get first 100 workbook files (any supported format)
            string[] workbookFiles = Directory.GetFiles(_sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            int processedCount = 0;

            foreach (string filePath in workbookFiles)
            {
                if (processedCount >= 100)
                    break;

                // Load workbook (using provided load rule)
                Workbook workbook = LoadWorkbook(filePath);

                // Add optional metadata if supplied
                if (_optionalMetadata.HasValue)
                {
                    // Use custom document property to store metadata
                    workbook.CustomDocumentProperties.Add(_optionalMetadata.Value.Key, _optionalMetadata.Value.Value);
                }

                // Remove styles that are not used in the workbook (using provided rule)
                workbook.RemoveUnusedStyles();

                // Determine output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(_outputFolder, fileName);

                // Save workbook (using provided save rule)
                SaveWorkbook(workbook, outputPath);

                processedCount++;
            }
        }

        // Load rule placeholder – loads a workbook from a file path
        private Workbook LoadWorkbook(string path)
        {
            // The actual rule implementation is assumed to be provided elsewhere.
            // Here we invoke the standard Aspose.Cells load method.
            return new Workbook(path);
        }

        // Save rule placeholder – saves a workbook to a file path
        private void SaveWorkbook(Workbook workbook, string path)
        {
            // The actual rule implementation is assumed to be provided elsewhere.
            // Here we invoke the standard Aspose.Cells save method.
            workbook.Save(path);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourceFolder = @"C:\Workbooks\Input";
            string outputFolder = @"C:\Workbooks\Processed";

            // Optional metadata example: add a custom property "ProcessedBy" with value "BatchJob"
            var metadata = ("ProcessedBy", "BatchJob");

            var processor = new WorkbookProcessor(sourceFolder, outputFolder, metadata);
            processor.ProcessAll();

            Console.WriteLine("Batch processing completed.");
        }
    }
}
