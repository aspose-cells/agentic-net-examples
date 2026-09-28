// Title: How to add a hidden worksheet with XML metadata and later retrieve it using Aspose.Cells for .NET
// AI Prompts: Using Aspose.Cells, create a new workbook, add a hidden worksheet named "Metadata", write an XML metadata string into cell A1, and save the workbook as an .xlsx file. | Load the saved .xlsx file with Aspose.Cells, locate the hidden "Metadata" worksheet, read the XML content from cell A1, and output it to the console.
// Common Searches: Aspose.Cells C# store XML string in hidden worksheet and read it later | how to hide a worksheet and embed custom XML in Aspose.Cells | retrieve XML metadata from a hidden sheet in an Excel file using Aspose.Cells .NET | save and load hidden worksheet containing XML data with Aspose.Cells | C# Aspose.Cells example for hidden sheet as metadata container
// Tags: add invisible sheet for metadata Aspose.Cells | store XML in cell A1 Aspose.Cells | extract XML from invisible sheet Aspose.Cells | save workbook in XLSX format Aspose.Cells | load workbook and access hidden sheet Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates creating a hidden "Metadata" worksheet, writing an XML metadata string to cell A1, saving the workbook as XLSX, then loading the file, locating the hidden sheet, and retrieving the XML content using Aspose.Cells for .NET.
class CustomXmlPartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a visible worksheet for data (optional)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Define custom XML metadata
            string xmlMetadata = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Metadata>
    <Author>John Doe</Author>
    <Created>2026-09-14</Created>
    <Description>Sample metadata for Aspose.Cells</Description>
</Metadata>";

            // -----------------------------------------------------------------
            // Store the XML metadata in a hidden worksheet (as a fallback for
            // environments where CustomXmlPart API is unavailable)
            // -----------------------------------------------------------------
            Worksheet metaSheet = workbook.Worksheets.Add("Metadata");
            metaSheet.IsVisible = false; // hide the sheet
            metaSheet.Cells["A1"].PutValue(xmlMetadata);

            // Save the workbook to a file
            string filePath = "CustomXmlWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // Later: Load the workbook and retrieve the stored XML
            // -------------------------------------------------
            if (File.Exists(filePath))
            {
                Workbook loadedWorkbook = new Workbook(filePath);

                // Find the hidden metadata worksheet
                Worksheet retrievedMetaSheet = null;
                foreach (Worksheet ws in loadedWorkbook.Worksheets)
                {
                    if (ws.Name.Equals("Metadata", StringComparison.OrdinalIgnoreCase))
                    {
                        retrievedMetaSheet = ws;
                        break;
                    }
                }

                if (retrievedMetaSheet != null)
                {
                    // Read the XML content from cell A1
                    string retrievedXml = retrievedMetaSheet.Cells["A1"].StringValue;
                    Console.WriteLine("Retrieved Custom XML Part:");
                    Console.WriteLine(retrievedXml);
                }
                else
                {
                    Console.WriteLine("Metadata worksheet not found.");
                }
            }
            else
            {
                Console.WriteLine($"File '{filePath}' was not found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
