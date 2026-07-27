using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsPropertyCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook (replace with actual file path)
                string sourcePath = "source.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook
                Workbook destinationWorkbook = new Workbook();

                // ----- Copy Built‑in Document Properties -----
                foreach (DocumentProperty srcProp in sourceWorkbook.BuiltInDocumentProperties)
                {
                    // Destination always contains the same built‑in properties
                    destinationWorkbook.BuiltInDocumentProperties[srcProp.Name].Value = srcProp.Value;
                }

                // ----- Copy Custom Document Properties -----
                foreach (DocumentProperty srcProp in sourceWorkbook.CustomDocumentProperties)
                {
                    if (destinationWorkbook.CustomDocumentProperties.Contains(srcProp.Name))
                    {
                        // Update existing property
                        destinationWorkbook.CustomDocumentProperties[srcProp.Name].Value = srcProp.Value;
                    }
                    else
                    {
                        // Add new property (value converted to string if required by overload)
                        destinationWorkbook.CustomDocumentProperties.Add(srcProp.Name, srcProp.Value?.ToString() ?? string.Empty);
                    }
                }

                // Save the destination workbook with the copied properties
                string destPath = "destination.xlsx";
                destinationWorkbook.Save(destPath);

                Console.WriteLine("Document properties copied successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}