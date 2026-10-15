// Title: Export an Aspose.Cells workbook to XML in C# – custom root element not supported
// Common Searches: how to save an Aspose.Cells workbook as XML using C# | Aspose.Cells XmlSaveOptions example for exporting Excel to XML | C# code to convert Excel workbook to XML with Aspose.Cells | limitations of custom root node when exporting XML with Aspose.Cells | Aspose.Cells export to XML without custom root element
// Tags: Aspose.Cells XML export C# | XmlSaveOptions usage Aspose.Cells | export workbook to XML Aspose.Cells | custom root element limitation Aspose.Cells | C# Excel to XML conversion Aspose

using System;
using Aspose.Cells;

// Demonstrates exporting a newly created Aspose.Cells workbook to an XML file in C#. The example uses XmlSaveOptions, and notes that Aspose.Cells does not allow setting a custom root element for the generated XML.
class ExportWorkbookAsXml
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Configure XML save options (custom root element name is not supported in Aspose.Cells)
            XmlSaveOptions xmlOptions = new XmlSaveOptions();

            // Export the workbook to XML using the specified options
            workbook.Save("ExportedWorkbook.xml", xmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
