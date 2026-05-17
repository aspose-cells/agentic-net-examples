using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ImportXmlMapFromXsdDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Import XML data into the first worksheet starting at cell A1
            // The XML file should conform to the required schema
            workbook.ImportXml("data.xml", "Sheet1", 0, 0);

            // Save the workbook with the imported data
            workbook.Save("output.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ImportXmlMapFromXsdDemo.Run();
        }
    }
}