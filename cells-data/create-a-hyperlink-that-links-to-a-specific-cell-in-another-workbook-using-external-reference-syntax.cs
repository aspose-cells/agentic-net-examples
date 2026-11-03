// Title: Add a hyperlink that points to a specific cell in another workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to cell in external workbook | how to use external reference syntax for hyperlinks in Aspose.Cells | create link to another workbook's sheet cell with Aspose.Cells .NET | C# Aspose.Cells Hyperlinks.Add external workbook example | reference "[Workbook]Sheet'!Cell" in Aspose.Cells hyperlink
// Tags: add external workbook hyperlink Aspose.Cells | Hyperlinks.Add external reference C# | Aspose.Cells external cell link syntax | C# create hyperlink to another workbook | Aspose.Cells workbook-to-workbook hyperlink

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes display text in cell A1, builds an external reference address using the "'[Workbook]Sheet'!Cell" syntax, adds a hyperlink to that address with Worksheet.Hyperlinks.Add, and saves the file as HyperlinkExternal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set display text for the hyperlink in cell A1
            worksheet.Cells["A1"].PutValue("Link to external cell");

            // Define the external reference address.
            // Syntax: '[WorkbookName]SheetName'!CellReference
            string externalAddress = "'[ExternalWorkbook.xlsx]Sheet1'!B2";

            // Add a hyperlink that points to the external cell.
            // The Add method returns the index of the created hyperlink.
            int hyperlinkIndex = worksheet.Hyperlinks.Add(0, 0, 0, 0, externalAddress);

            // (Optional) Retrieve the hyperlink object if further customization is needed
            // Hyperlink hyperlink = worksheet.Hyperlinks[hyperlinkIndex];

            // Save the workbook
            string outputPath = "HyperlinkExternal.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
