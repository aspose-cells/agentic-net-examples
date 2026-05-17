using System;
using Aspose.Cells;

namespace AsposeCellsXmlValidation
{
    public class ValidateLinkedCells
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Sample XML data to be imported
            string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Data>
        <Item>Value1</Item>
    </Data>
</Root>";

            // Import the XML data into the first worksheet starting at cell A1
            workbook.ImportXml(xmlData, "Sheet1", 0, 0);

            // Access the first worksheet (the one we imported into)
            Worksheet worksheet = workbook.Worksheets[0];

            // Expected value for the XML element at the given path
            const string expectedValue = "Value1";

            // The imported XML places the value in cell A1
            string cellValue = worksheet.Cells["A1"].StringValue;

            bool isValid = string.Equals(cellValue, expectedValue, StringComparison.Ordinal);
            Console.WriteLine($"Cell A1 linked to '/Root/Data/Item': Value = '{cellValue}' => {(isValid ? "Valid" : "Invalid")}");

            // Save the workbook (optional, just to demonstrate lifecycle compliance)
            workbook.Save("ValidatedXmlWorkbook.xlsx");
        }
    }
}