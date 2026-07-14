using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsLinkStatusDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the built‑in document properties collection
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

            // Retrieve the LinksUpToDate property which indicates whether hyperlinks are up‑to‑date
            bool linksAreUpToDate = builtInProps.LinksUpToDate;

            // Output the status to the console
            Console.WriteLine($"Links up‑to‑date: {linksAreUpToDate}");

            // (Optional) List all hyperlinks in the first worksheet for reference
            Worksheet sheet = workbook.Worksheets[0];
            foreach (Hyperlink link in sheet.Hyperlinks)
            {
                Console.WriteLine($"Hyperlink address: {link.Address}");
            }
        }
    }
}