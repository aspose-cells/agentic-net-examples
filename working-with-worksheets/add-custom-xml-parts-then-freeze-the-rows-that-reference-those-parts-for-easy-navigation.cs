// Title: Add a custom XML part to an Excel workbook and freeze header rows using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add custom XML part to workbook and retrieve its ID | Freeze top rows in Excel with Aspose.Cells after inserting custom XML data | How to reference a custom XML part ID in worksheet cells using Aspose.Cells for .NET | C# code to add XML part to workbook and apply FreezePanes at a specific cell | Saving a workbook with a custom XML part and frozen panes using Aspose.Cells
// Tags: add custom xml part Aspose.Cells C# | freeze header rows Aspose.Cells FreezePanes | worksheet cell linking to custom xml part id | save workbook with custom xml part Aspose.Cells | populate worksheet with xml data using Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace CustomXmlPartAndFreezeDemo
{
    // The example creates a new workbook, adds a custom XML part containing product information, assigns a GUID as the part ID, writes product rows together with the XML part ID into the worksheet, freezes the first two rows to keep the reference column visible, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // 2. Prepare sample XML data (no schema required for this demo).
                string xmlData = @"<Products xmlns=""http://example.com/products"">
                                      <Product>
                                          <ID>1</ID>
                                          <Name>Apple</Name>
                                          <Price>0.5</Price>
                                      </Product>
                                      <Product>
                                          <ID>2</ID>
                                          <Name>Banana</Name>
                                          <Price>0.3</Price>
                                      </Product>
                                   </Products>";
                byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
                byte[] schemaBytes = null; // No schema

                // 3. Add the custom XML part to the workbook.
                int partIndex = workbook.CustomXmlParts.Add(dataBytes, schemaBytes);
                CustomXmlPart xmlPart = workbook.CustomXmlParts[partIndex];

                // Assign a readable GUID as the part ID.
                string partId = Guid.NewGuid().ToString();
                xmlPart.ID = partId;
                Console.WriteLine($"Custom XML Part added with ID: {partId}");

                // 4. Populate the worksheet with data and a reference to the XML part.
                sheet.Cells["A1"].PutValue("Product ID");
                sheet.Cells["B1"].PutValue("Product Name");
                sheet.Cells["C1"].PutValue("Price");
                sheet.Cells["D1"].PutValue("XML Part ID"); // Reference column

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(0.5);
                sheet.Cells["D2"].PutValue(partId);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(0.3);
                sheet.Cells["D3"].PutValue(partId);

                // 5. Freeze the first two rows so the reference column stays visible.
                // Freeze at cell A3 with 2 frozen rows and 0 frozen columns.
                sheet.FreezePanes("A3", 2, 0);
                Console.WriteLine("Rows 1‑2 have been frozen for easy navigation.");

                // 6. Save the workbook.
                string outputPath = "CustomXmlPart_FreezeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
