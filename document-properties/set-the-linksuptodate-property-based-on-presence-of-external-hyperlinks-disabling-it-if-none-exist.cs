using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class SetLinksUpToDateBasedOnHyperlinks
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: add a hyperlink (comment out to test the "no hyperlinks" case)
            // sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Determine whether any worksheet contains at least one hyperlink
            bool hasHyperlinks = false;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Hyperlinks.Count > 0)
                {
                    hasHyperlinks = true;
                    break;
                }
            }

            // Set the LinksUpToDate property based on the presence of hyperlinks
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;
            builtInProps.LinksUpToDate = hasHyperlinks;

            // Define output file path
            string outputPath = "LinksUpToDateResult.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}