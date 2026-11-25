// Title: Import an XML file as an XmlMap into an Excel workbook using Aspose.Cells Workbook.ImportXml (C#)
// Common Searches: C# Aspose.Cells import XML file into worksheet with Workbook.ImportXml | How to create an XmlMap from an external XML file using Aspose.Cells | Retrieve XmlMap name and root element after importing XML in Aspose.Cells | Save workbook as XLSX after importing XML data with Aspose.Cells .NET | Delete temporary XML file after importing it into an Excel workbook using Aspose.Cells
// Tags: Workbook.ImportXml XML to Excel conversion | Aspose.Cells create XmlMap from file | import XML data into worksheet C# | access XmlMap properties Aspose.Cells | save workbook as XLSX after XML import

using System;
using System.IO;
using Aspose.Cells;

namespace ImportXmlMapDemo
{
    // The example creates a temporary XML file containing product data, uses Aspose.Cells Workbook.ImportXml to import the XML into the first worksheet at cell A1 (automatically generating an XmlMap), displays the XmlMap's name and root element, prints the imported cell values, saves the workbook as an XLSX file, and finally removes the temporary XML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file names
            const string xmlFilePath = "SampleData.xml";
            const string outputWorkbookPath = "ImportedXmlWorkbook.xlsx";

            // Prepare a simple XML file that matches a typical schema
            // The XML structure will be:
            // <Products>
            //   <Product>
            //     <Name>Laptop</Name>
            //     <Price>999.99</Price>
            //   </Product>
            //   <Product>
            //     <Name>Phone</Name>
            //     <Price>699.99</Price>
            //   </Product>
            // </Products>
            string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Products>
    <Product>
        <Name>Laptop</Name>
        <Price>999.99</Price>
    </Product>
    <Product>
        <Name>Phone</Name>
        <Price>699.99</Price>
    </Product>
</Products>";

            try
            {
                // Write the XML content to a file
                File.WriteAllText(xmlFilePath, xmlContent);
                Console.WriteLine($"XML file '{xmlFilePath}' created.");

                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Import the XML data into the first worksheet starting at cell A1 (row 0, column 0)
                // This also creates an XmlMap internally.
                workbook.ImportXml(xmlFilePath, "Sheet1", 0, 0);
                Console.WriteLine("XML data imported into workbook.");

                // Verify that an XmlMap was created
                if (workbook.Worksheets.XmlMaps.Count > 0)
                {
                    XmlMap map = workbook.Worksheets.XmlMaps[0];
                    Console.WriteLine($"XmlMap created with name: {map.Name}");
                    Console.WriteLine($"Root element of the map: {map.RootElementName}");
                }
                else
                {
                    Console.WriteLine("No XmlMap was created after ImportXml.");
                }

                // Optionally, display a few imported cell values to the console
                Worksheet sheet = workbook.Worksheets["Sheet1"];
                Console.WriteLine("Imported cell values:");
                Console.WriteLine($"A1: {sheet.Cells["A1"].StringValue}");
                Console.WriteLine($"B1: {sheet.Cells["B1"].StringValue}");
                Console.WriteLine($"A2: {sheet.Cells["A2"].StringValue}");
                Console.WriteLine($"B2: {sheet.Cells["B2"].StringValue}");
                Console.WriteLine($"A3: {sheet.Cells["A3"].StringValue}");
                Console.WriteLine($"B3: {sheet.Cells["B3"].StringValue}");

                // Save the workbook to disk
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as '{outputWorkbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up the temporary XML file
                try
                {
                    if (File.Exists(xmlFilePath))
                    {
                        File.Delete(xmlFilePath);
                        Console.WriteLine($"Temporary XML file '{xmlFilePath}' deleted.");
                    }
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine($"Failed to delete temporary XML file: {cleanupEx.Message}");
                }
            }
        }
    }
}
