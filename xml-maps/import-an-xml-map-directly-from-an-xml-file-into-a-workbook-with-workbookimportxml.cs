using System;
using Aspose.Cells;

namespace AsposeCellsXmlImportDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Import XML data from a file into the first worksheet starting at cell A1 (row 0, column 0)
            // Parameters: xml file path, destination sheet name, start row index, start column index
            workbook.ImportXml("data.xml", "Sheet1", 0, 0);

            // Save the workbook to an Excel file
            workbook.Save("output.xlsx");

            Console.WriteLine("XML data imported and workbook saved successfully.");
        }
    }
}