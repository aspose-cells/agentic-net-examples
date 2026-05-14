using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LinksUpToDateDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: add a hyperlink (you can comment this out to test the "no links" case)
            // The hyperlink points to an external URL, which counts as an external link
            sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Determine if the workbook contains any external links (hyperlinks or external formulas)
            bool hasExternalLinks = workbook.Worksheets.ExternalLinks.Count > 0 || sheet.Hyperlinks.Count > 0;

            // Set the LinksUpToDate property accordingly
            workbook.BuiltInDocumentProperties.LinksUpToDate = hasExternalLinks;

            // Save the workbook
            workbook.Save("LinksUpToDateResult.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LinksUpToDateDemo.Run();
        }
    }
}