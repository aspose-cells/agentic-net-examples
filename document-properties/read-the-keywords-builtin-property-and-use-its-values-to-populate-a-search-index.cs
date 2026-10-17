// Title: Extract the Keywords built‑in document property from an Excel workbook and create a searchable keyword list with Aspose.Cells for .NET
// Common Searches: how to get the Keywords built‑in property from an Excel workbook using Aspose.Cells in C# | reading and splitting Excel Keywords document property into a List<string> with Aspose.Cells | populate a keyword collection from Excel metadata for search indexing in .NET | handle missing Keywords property when reading built‑in document properties with Aspose.Cells | convert comma‑separated Keywords property to a searchable list in C#
// Tags: read built‑in properties with Aspose.Cells | extract Excel Keywords metadata in C# | parse comma‑separated keywords into list | build keyword search index from workbook metadata | gracefully handle absent document property

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook using Aspose.Cells, accesses its BuiltInDocumentProperties collection, retrieves the "Keywords" property (or an empty string if the property is missing), splits the comma‑separated keywords, trims each entry, adds them to a List<string> to serve as a simple search index, and writes each term to the console. An optional workbook save line is provided but commented out.
class Program
{
    static void Main()
    {
        // Load the Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        var builtInProps = workbook.BuiltInDocumentProperties;

        // Retrieve the value of the "Keywords" property; handle missing property gracefully
        string keywords = builtInProps["Keywords"]?.ToString() ?? string.Empty;

        // Prepare a simple search index (list of individual keyword strings)
        List<string> searchIndex = new List<string>();

        // Split the Keywords string by commas (common delimiter) and trim whitespace
        foreach (string keyword in keywords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
        {
            searchIndex.Add(keyword.Trim());
        }

        // Example usage: display the populated search index
        foreach (string term in searchIndex)
        {
            Console.WriteLine(term);
        }

        // If you need to save changes to the workbook, uncomment the following line
        // workbook.Save("output.xlsx");
    }
}
