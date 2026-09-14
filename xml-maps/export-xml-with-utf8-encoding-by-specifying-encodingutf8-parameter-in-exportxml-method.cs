// Title: Export an Excel workbook to a UTF-8 encoded XML file using Aspose.Cells ExportXml in C#
// AI Prompts: Generate C# code that creates a Workbook, populates cells, and calls Workbook.ExportXml with Encoding.UTF8.WebName to produce a UTF-8 XML file. | Show how to wrap the ExportXml call in a try‑catch block to handle errors when saving an Excel workbook as UTF-8 XML with Aspose.Cells. | Explain the steps required to set the encoding argument for ExportXml so the resulting XML uses UTF-8 character encoding.
// Common Searches: Aspose.Cells ExportXml method specify UTF-8 encoding in C# | How to export an Excel workbook to XML with UTF-8 using Aspose.Cells library | C# sample code for Workbook.ExportXml with custom encoding parameter | Saving Excel as UTF-8 XML file with Aspose.Cells | ExportXml encoding parameter example for UTF-8 in .NET
// Tags: Aspose.Cells ExportXml UTF-8 encoding | C# export workbook to XML with custom encoding | Workbook.ExportXml encoding parameter | UTF-8 XML output Aspose.Cells | Excel to XML conversion C# Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;

// The example creates a new Workbook, adds header and data cells, and then exports the workbook to an XML file named "output.xml" using Aspose.Cells' ExportXml method with the encoding set to UTF-8 via Encoding.UTF8.WebName. The operation is enclosed in a try‑catch block to capture any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);

            // Export the workbook to an XML file using UTF-8 encoding
            // ExportXml expects the encoding name as a string (e.g., "UTF-8")
            workbook.ExportXml("output.xml", Encoding.UTF8.WebName);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
