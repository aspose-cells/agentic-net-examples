// Title: Import XML from a Stream into an Aspose.Cells Workbook with Linked Cells (C#)
// Description: Demonstrates how to create a Workbook, convert an XML string to a MemoryStream, and use the ImportXml overload to load the XML into a specified worksheet starting at cell A1, preserving the XML map, then save the file as an Excel workbook.
// Keywords: Aspose.Cells ImportXml stream C# | load XML into Excel workbook | MemoryStream XML import Aspose | linked cells XML map Aspose.Cells | ImportXml overload example | C# Excel XML data import | Aspose.Cells XML map refresh
// Common Searches: Aspose.Cells ImportXml from MemoryStream example | How to import XML into a worksheet using C# | Create linked cells from XML with Aspose.Cells | Save workbook after ImportXml stream | ImportXml overload parameters
// Developer Intent: Load XML content supplied via a stream into a new workbook, map it to cells, and generate an Excel file.
// Use Cases: Transform XML configuration received over a network into an Excel report. | Populate a spreadsheet with product data stored in an XML string for analysis. | Maintain a live connection between worksheet cells and source XML for periodic refreshes.
// AI Prompts: Generate C# code that reads XML from a MemoryStream and imports it into a specific sheet using Aspose.Cells ImportXml while keeping the XML map intact. | Show an example of handling large XML streams with ImportXml overload and saving the workbook efficiently. | Explain best practices for error handling and resource cleanup when using ImportXml with a stream in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a Workbook, convert an XML string to a MemoryStream, and use the ImportXml overload to load the XML into a specified worksheet starting at cell A1, preserving the XML map, then save the file as an Excel workbook.
    public class ImportXmlFromStreamDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // Sample XML content to be imported
                string xmlData = @"<Products>
    <Product>
        <Name>Laptop</Name>
        <Price>999.99</Price>
    </Product>
    <Product>
        <Name>Phone</Name>
        <Price>699.99</Price>
    </Product>
</Products>";

                // Convert the XML string into a MemoryStream
                using (MemoryStream xmlStream = new MemoryStream())
                {
                    using (StreamWriter writer = new StreamWriter(xmlStream))
                    {
                        writer.Write(xmlData);
                        writer.Flush();
                        xmlStream.Position = 0; // Reset position for reading

                        // Import the XML data into the first worksheet at cell A1 (row 0, column 0)
                        workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
                    }
                }

                // Save the workbook to an Excel file
                string outputPath = "ImportXmlFromStreamDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ImportXmlFromStreamDemo.Run();
        }
    }
}
