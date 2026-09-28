// Title: How to read the Language built-in document property (locale) from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells and prints the workbook's BuiltInDocumentProperties.Language value. | Create a reusable C# method that accepts a file path or stream and returns the locale string from the workbook's Language built-in property using Aspose.Cells. | Show how to modify the example to retrieve the Language property from a workbook loaded from a MemoryStream instead of a file path.
// Common Searches: aspnet read language built-in property from excel file using aspose.cells | c# get locale setting from workbook builtindocumentproperties language | how to extract excel document language (locale) with aspose.cells in .net | retrieve language built-in property from .xlsx using aspose.cells c# example | asp.net core read workbook language property aspose.cells
// Tags: language built-in property extraction Aspose.Cells | workbook locale retrieval Aspose.Cells C# | builtindocumentproperties language access Aspose.Cells | xlsx language attribute reading Aspose.Cells | c# get excel document locale Aspose.Cells

using System;
using Aspose.Cells;

// Loads 'input.xlsx' with Aspose.Cells, accesses the workbook's BuiltInDocumentProperties, reads the Language property (locale), and writes the value to the console.
class Program
{
    static void Main()
    {
        // Load the spreadsheet from a file
        var workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties
        var builtInProps = workbook.BuiltInDocumentProperties;

        // Read the Language property which indicates the locale settings
        string language = builtInProps.Language;

        // Output the language (locale) information
        Console.WriteLine("Document Language (Locale): " + language);
    }
}
