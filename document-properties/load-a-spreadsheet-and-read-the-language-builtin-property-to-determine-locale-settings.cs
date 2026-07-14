using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string inputPath = "sample.xlsx";

        // Create LoadOptions instance (uses provided constructor rule)
        LoadOptions loadOptions = new LoadOptions();

        // Load the workbook with the specified LoadOptions (uses provided load rule)
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Read the Language property which indicates the locale of the document
        string language = builtInProps.Language;

        Console.WriteLine($"Document Language: {language}");

        // Optional: save the workbook to demonstrate the save rule (can be omitted if not needed)
        workbook.Save("output.xlsx");
    }
}