using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the collection of built‑in document properties
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Inspect the LinksUpToDate property which indicates if hyperlinks are current
        bool linksUpToDate = builtInProps.LinksUpToDate;

        // Output the status
        Console.WriteLine("Links up to date: " + linksUpToDate);
    }
}