// Title: Read the DocumentVersion built-in property from an Aspose.Cells Workbook using C#
// Common Searches: C# Aspose.Cells how to read the DocumentVersion built-in property from a workbook | Get Excel document version property with Aspose.Cells in .NET | Retrieve built-in DocumentVersion from a new Workbook using Aspose.Cells for C# | Aspose.Cells read built-in properties like DocumentVersion programmatically | Example code to access DocumentVersion property in Aspose.Cells .NET
// Tags: Aspose.Cells read built-in document version | C# retrieve Excel DocumentVersion property | Workbook built-in properties extraction Aspose | Aspose.Cells DocumentVersion access example | C# get built-in properties from Excel workbook | Aspose.Cells .NET built-in property usage

using Aspose.Cells;
using System;

// The example creates a Workbook, accesses its BuiltInDocumentProperties collection, reads the "DocumentVersion" property (or reports it as not set), and prints the value to the console.
class Program
{
    static void Main()
    {
        // Instantiate a new workbook
        Workbook workbook = new Workbook();

        // Retrieve the built‑in property "DocumentVersion"
        var docVersionProp = workbook.BuiltInDocumentProperties["DocumentVersion"];
        string version = docVersionProp != null ? docVersionProp.Value?.ToString() : "Not set";

        // Display the version information
        Console.WriteLine($"DocumentVersion: {version}");
    }
}
