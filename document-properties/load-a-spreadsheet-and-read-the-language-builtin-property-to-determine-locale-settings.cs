using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
            return;
        }

        try
        {
            // Create LoadOptions (default constructor)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the built‑in document properties collection
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

            // Retrieve the Language property which indicates the locale settings
            string language = builtInProps.Language;

            // Display the language (locale) information
            Console.WriteLine("Document Language: " + language);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}