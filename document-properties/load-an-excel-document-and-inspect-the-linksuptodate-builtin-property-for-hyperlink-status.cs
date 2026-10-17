// Title: Read the LinksUpToDate built‑in document property from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to get the LinksUpToDate built‑in property from an Excel file | Check if external hyperlinks are up to date in an .xlsx using Aspose.Cells | Retrieve workbook built‑in document properties in C# with Aspose.Cells | Determine whether Excel links are current via Aspose.Cells API | C# code to read LinksUpToDate flag from an Excel workbook
// Tags: Aspose.Cells read LinksUpToDate property | C# get Excel built-in document properties | check external link status Aspose.Cells | retrieve hyperlink freshness Excel .NET | load workbook inspect link up‑to‑date flag

using Aspose.Cells;
using System;

// The example loads an Excel workbook with Aspose.Cells, accesses the BuiltInDocumentProperties.LinksUpToDate boolean to determine if all external links are current, and writes the result to the console.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the built‑in property that indicates whether all external links are up to date
        bool linksUpToDate = workbook.BuiltInDocumentProperties.LinksUpToDate;

        // Output the status
        Console.WriteLine($"LinksUpToDate: {linksUpToDate}");
    }
}
