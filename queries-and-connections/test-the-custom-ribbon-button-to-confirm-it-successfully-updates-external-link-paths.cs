using System;
using System.IO;
using Aspose.Cells;

namespace RibbonButtonExternalLinkTest
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and add a formula that contains an external link.
                // Use a local file path format that Aspose.Cells accepts for external references.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                // Example external link (adjust the path as needed for your environment)
                sheet.Cells["A1"].Formula = "='C:\\Temp\\SourceFile.xlsx'!Sheet1!A1";

                // 2. Simulate the custom ribbon button action:
                //    Update all external link paths to a new location.
                for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                {
                    ExternalLink link = workbook.Worksheets.ExternalLinks[i];
                    string original = link.OriginalDataSource;

                    // Replace the old base folder with the new one
                    string updated = original.Replace(
                        @"C:\Temp\",
                        @"D:\SharedDocuments\");

                    // Apply the updated path back to the external link
                    link.OriginalDataSource = updated;
                }

                // 3. Verify the changes by writing the modified paths to the console
                foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
                {
                    Console.WriteLine("Modified External Link Path: " + link.OriginalDataSource);
                }

                // 4. Save the workbook (the ribbon button would normally trigger this)
                string outputPath = "RibbonButtonExternalLinkTest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}