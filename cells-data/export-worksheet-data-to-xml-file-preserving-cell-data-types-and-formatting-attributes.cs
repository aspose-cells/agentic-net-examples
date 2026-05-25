using System;
using Aspose.Cells;

namespace ExportWorksheetToXml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate header row
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["C1"].PutValue("IsActive");

            // Apply date format to cell A2
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // m/d/yyyy
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A2"].SetStyle(dateStyle);

            // Apply currency format to cell B2
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 164; // $#,##0.00
            sheet.Cells["B2"].PutValue(1234.56);
            sheet.Cells["B2"].SetStyle(currencyStyle);

            // Boolean value in cell C2
            sheet.Cells["C2"].PutValue(true);

            // Configure XmlSaveOptions to preserve data types and formatting
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                SheetNameAsElementName = true, // use sheet name as XML element name
                DataAsAttribute = false,       // keep data as element values
                HasHeaderRow = true            // first row is treated as header
                // SheetIndexes left null to export all sheets
            };

            // Save the workbook as an XML file
            workbook.Save("ExportedData.xml", saveOptions);
        }
    }
}