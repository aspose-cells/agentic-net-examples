// Title: Read the LinksUpToDate built‑in document property from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells and returns the value of the LinksUpToDate built‑in property. | Create a reusable C# method that accepts a file path, loads the workbook with Aspose.Cells, and prints whether all external links are up to date. | Show how to safely access and cast the LinksUpToDate built‑in property to a boolean, handling cases where the property is missing.
// Common Searches: Aspose.Cells C# read LinksUpToDate built‑in document property from Excel file | How to check if external links in an .xlsx are up to date using Aspose.Cells .NET | Retrieve boolean LinksUpToDate property from workbook programmatically with Aspose.Cells | C# Aspose.Cells example to get hyperlink status via LinksUpToDate property | Determine if Excel workbook links are current using Aspose.Cells built‑in properties
// Tags: Aspose.Cells read built-in document property | C# retrieve LinksUpToDate property | Excel workbook external links status Aspose.Cells | load .xlsx inspect LinksUpToDate | boolean built-in property Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, accesses the built‑in LinksUpToDate property, safely casts it to a boolean, and writes the up‑to‑date status of external links to the console.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the built‑in property "LinksUpToDate"
        // This property indicates whether all external links (including hyperlinks) are up to date.
        bool linksUpToDate = false;
        var prop = workbook.BuiltInDocumentProperties["LinksUpToDate"];
        if (prop != null && prop.Value != null)
        {
            // The property value is stored as a boolean
            linksUpToDate = (bool)prop.Value;
        }

        // Output the status of the LinksUpToDate property
        Console.WriteLine($"LinksUpToDate: {linksUpToDate}");
    }
}
