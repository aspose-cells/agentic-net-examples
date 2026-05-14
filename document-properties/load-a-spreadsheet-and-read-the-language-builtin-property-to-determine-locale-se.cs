using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create default load options (uses provided LoadOptions() rule)
        LoadOptions loadOptions = new LoadOptions();

        // Load the workbook with the specified options (uses Workbook(string, LoadOptions) rule)
        Workbook workbook = new Workbook("sample.xlsx", loadOptions);

        // Read the built‑in Language property from the document properties
        string language = workbook.BuiltInDocumentProperties.Language;

        // Display the language (locale) information
        Console.WriteLine($"Document Language Property: {language}");
    }
}