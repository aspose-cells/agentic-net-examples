using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsXmlMapFromString
{
    class Program
    {
        static void Main()
        {
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

            Workbook workbook = new Workbook();

            // Import the XML data into the first worksheet starting at cell A1
            using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlData)))
            {
                workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
            }

            workbook.Save("ProductsMapped.xlsx");
        }
    }
}