// Title: How to read the Language built‑in document property of an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read Language built‑in document property from .xlsx | Get locale setting of an Excel file using Aspose.Cells in .NET | How to access workbook language metadata with Aspose.Cells C# example | Retrieve Excel workbook built‑in Language property programmatically | Aspose.Cells read workbook language property for localization
// Tags: read built-in language property Aspose.Cells | access workbook locale metadata C# | Aspose.Cells get Excel language setting | C# retrieve Excel built-in document property | Aspose.Cells document properties example

using System;
using Aspose.Cells;

// The sample loads an Excel file (input.xlsx) into an Aspose.Cells Workbook, accesses the built‑in document property named "Language", checks for its presence, and prints the locale value to the console or reports that the property is missing.
class Program
{
    static void Main()
    {
        // Load the spreadsheet from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document property named "Language"
        // This property stores the locale setting of the workbook
        var languageProperty = workbook.BuiltInDocumentProperties["Language"];

        if (languageProperty != null && languageProperty.Value != null)
        {
            // Output the language/locale value
            Console.WriteLine($"Language (locale) setting: {languageProperty.Value}");
        }
        else
        {
            Console.WriteLine("The Language property is not set in this workbook.");
        }
    }
}
