// Title: Batch add a shared ContentTypeProperty to multiple Excel workbooks using Aspose.Cells for .NET
// Description: C# program that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, inserts a ContentTypeProperty named "SharedProperty" with value "SharedValue", saves the file in place, and logs success or error messages. Includes directory validation and exception handling.
// Keywords: Aspose.Cells | ContentTypeProperty | C# batch update Excel | add custom property to workbooks | process multiple .xlsx files | Excel metadata automation | .NET Excel library
// Common Searches: how to add the same ContentTypeProperty to all Excel files with Aspose.Cells | C# loop through folder of .xlsx files and set a custom property | batch add metadata to multiple workbooks using Aspose.Cells | Aspose.Cells add custom property to each workbook in a directory
// Developer Intent: Insert an identical ContentTypeProperty into every workbook located in a specified directory.
// Use Cases: Embed a company‑wide tag (e.g., department code) into all report workbooks before distribution. | Update a version or release identifier across a set of template files in one operation. | Apply a compliance disclaimer property to every workbook to satisfy regulatory requirements.
// AI Prompts: Write C# code that uses Aspose.Cells to add a shared ContentTypeProperty to all .xlsx files in a given folder, with checks for missing directories and file‑level errors. | Show how to verify whether a ContentTypeProperty already exists in a workbook before adding it, to prevent duplicate entries.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // C# program that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, inserts a ContentTypeProperty named "SharedProperty" with value "SharedValue", saves the file in place, and logs success or error messages. Includes directory validation and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the Excel workbooks
            string folderPath = @"C:\Workbooks";

            // Ensure the directory exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Directory not found: {folderPath}");
                return;
            }

            // Define the shared content type property name and value
            const string propertyName = "SharedProperty";
            const string propertyValue = "SharedValue";

            // Get all Excel files in the directory (adjust the pattern if needed)
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx");

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Add the shared content type property
                    // If the property already exists, this will add another entry with the same name
                    workbook.ContentTypeProperties.Add(propertyName, propertyValue);

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath);

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
