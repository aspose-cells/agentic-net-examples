using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        try
        {
            // Paths to the source and destination workbooks
            string sourcePath = "source.xlsx";
            string destPath = "destination.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook
            Workbook destWorkbook = new Workbook();

            // ----- Copy Built‑in Document Properties -----
            foreach (DocumentProperty srcProp in sourceWorkbook.BuiltInDocumentProperties)
            {
                // Preserve the same value in the destination workbook
                destWorkbook.BuiltInDocumentProperties[srcProp.Name].Value = srcProp.Value;
            }

            // ----- Copy Custom Document Properties -----
            foreach (DocumentProperty srcProp in sourceWorkbook.CustomDocumentProperties)
            {
                try
                {
                    // Update existing property
                    DocumentProperty destProp = destWorkbook.CustomDocumentProperties[srcProp.Name];
                    destProp.Value = srcProp.Value;
                }
                catch (ArgumentException)
                {
                    // Property does not exist; add it (convert value to string if necessary)
                    string valueAsString = srcProp.Value?.ToString() ?? string.Empty;
                    destWorkbook.CustomDocumentProperties.Add(srcProp.Name, valueAsString);
                }
            }

            // Save the destination workbook
            destWorkbook.Save(destPath);
            Console.WriteLine($"Properties copied successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}