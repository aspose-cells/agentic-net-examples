// Title: Read the LinksUpToDate built‑in property (hyperlink status) from an Excel workbook using Aspose.Cells for .NET
// Description: This example shows how to open an Excel file with Aspose.Cells, access the BuiltInDocumentPropertyCollection, retrieve the boolean LinksUpToDate property that indicates whether external hyperlinks are current, and output the result.
// Keywords: Aspose.Cells LinksUpToDate | C# read Excel built‑in document property | hyperlink status Excel Aspose | check external links Aspose.Cells | BuiltInDocumentPropertyCollection LinksUpToDate | Aspose.Cells .NET document properties
// Common Searches: Aspose.Cells get LinksUpToDate property | C# read Excel hyperlink status built‑in property | How to verify workbook links are up to date using Aspose | Retrieve LinksUpToDate from Excel file in .NET | Excel hyperlink validation with Aspose.Cells
// Developer Intent: Load an Excel workbook and determine if its external hyperlinks are up to date by reading the LinksUpToDate built‑in property.
// Use Cases: Confirm that all external references in a generated report are current before distribution. | Add an automated quality‑check in a CI pipeline that flags workbooks with outdated hyperlinks. | Run a batch scan of multiple Excel files to log those where LinksUpToDate is false.
// AI Prompts: Generate C# code using Aspose.Cells that opens an Excel workbook and prints the LinksUpToDate property. | Create a snippet that logs a warning when LinksUpToDate is false and lists the workbook's external hyperlinks. | Write a method that returns the LinksUpToDate value and handles cases where the property is missing or inaccessible.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This example shows how to open an Excel file with Aspose.Cells, access the BuiltInDocumentPropertyCollection, retrieve the boolean LinksUpToDate property that indicates whether external hyperlinks are current, and output the result.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Retrieve the LinksUpToDate property which indicates hyperlink status
        bool linksUpToDate = builtInProps.LinksUpToDate;

        // Display the result
        Console.WriteLine($"Links up to date: {linksUpToDate}");
    }
}
