using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExternalLinkRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "LocalSheet";

            // Add an external link to another workbook (external.xlsx) and its sheet (Sheet1)
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            int linkIndex = externalLinks.Add("external.xlsx", new string[] { "Sheet1" });

            // Verify that the external link was added
            ExternalLink extLink = externalLinks[linkIndex];
            Console.WriteLine($"External link added. DataSource: {extLink.DataSource}");

            // Set a formula in cell A1 that references the external workbook
            sheet.Cells["A1"].Formula = $"='[{extLink.DataSource}]Sheet1'!A1";

            // Create a range that includes the cell with the external reference
            AsposeRange range = sheet.Cells.CreateRange("A1:B2");
            range.Name = "MyRangeWithExternalRef";

            // Output range details
            Console.WriteLine($"Range Name: {range.Name}");
            Console.WriteLine($"Range Address: {range.Address}");
            Console.WriteLine($"First Cell Formula: {range[0, 0].Formula}");

            // Save the workbook (XLSX format)
            workbook.Save("ExternalLinkRangeDemo.xlsx");
            Console.WriteLine("Workbook saved as ExternalLinkRangeDemo.xlsx");
        }
    }
}