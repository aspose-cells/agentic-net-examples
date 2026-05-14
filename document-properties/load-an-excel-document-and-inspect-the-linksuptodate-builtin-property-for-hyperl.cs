using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class InspectLinksUpToDate
{
    static void Main()
    {
        // Load an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Retrieve the LinksUpToDate property which indicates if hyperlinks are current
        bool linksAreUpToDate = builtInProps.LinksUpToDate;

        // Output the status
        Console.WriteLine("Hyperlinks up‑to‑date: " + linksAreUpToDate);
    }
}