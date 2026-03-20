using System;
using Aspose.Cells;

namespace AsposeCellsLoadSpreadsheetML
{
    class Program
    {
        static void Main()
        {
            // Path to the SpreadsheetML (Excel 2003 XML) file
            string xmlFilePath = "sample.xml";

            // Load the workbook (auto-detect format)
            Workbook workbook = new Workbook(xmlFilePath);

            // Example manipulation: read the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Further processing can be done here...
        }
    }
}