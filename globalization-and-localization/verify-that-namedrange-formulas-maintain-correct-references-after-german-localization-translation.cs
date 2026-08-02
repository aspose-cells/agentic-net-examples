// Title: C# – Verify Named Range References Remain Correct After German Localization in Aspose.Cells
// Description: Creates a workbook, populates cells A1‑A3, sets the workbook region to Germany, adds a named range "MyRange", retrieves its reference in both English and German using GetRefersTo, resolves the range with GetRange to confirm it points to A1:A3, and saves the file. Demonstrates how to validate formula localization for named ranges in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | named range | localization | German | GetRefersTo | GetRange | formula translation | CountryCode.Germany | region setting
// Common Searches: Aspose.Cells get German formula reference | named range localization German Aspose.Cells | verify named range after setting workbook region | C# Aspose.Cells GetRefersTo localized format | handle formula translation in Aspose.Cells
// Developer Intent: Confirm that a named range’s formula address stays accurate when the workbook is localized to German.
// Use Cases: Compare the English and German RefersTo strings of a named range to detect translation issues. | Assert that Name.GetRange resolves to the expected cells (A1:A3) after changing workbook.Settings.Region. | Save a localized workbook and verify that the named range persists with correct references. | Integrate automated checks for formula localization in multi‑language deployment pipelines.
// AI Prompts: Generate C# unit tests that compare the standard and German RefersTo values of a named range in Aspose.Cells. | Show code to programmatically validate that a named range resolves to A1:A3 after setting workbook.Settings.Region to CountryCode.Germany. | Explain best practices for managing formula localization of named ranges when exporting workbooks to different language regions using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeLocalizationVerification
{
    // Creates a workbook, populates cells A1‑A3, sets the workbook region to Germany, adds a named range "MyRange", retrieves its reference in both English and German using GetRefersTo, resolves the range with GetRange to confirm it points to A1:A3, and saves the file. Demonstrates how to validate formula localization for named ranges in Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate some data that the named range will refer to
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Set the workbook region to German to enable German localization
                workbook.Settings.Region = CountryCode.Germany;

                // Create a named range that refers to A1:A3 on Sheet1
                // In Aspose.Cells, named ranges are managed via the NameCollection of the workbook's worksheets
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                namedRange.RefersTo = "=Sheet1!$A$1:$A$3";

                // Retrieve the reference in standard (English) format
                string refersToStandard = namedRange.GetRefersTo(false, false);
                // Retrieve the reference in localized (German) format
                string refersToLocal = namedRange.GetRefersTo(false, true);

                Console.WriteLine("Reference in standard format: " + refersToStandard);
                Console.WriteLine("Reference in German localized format: " + refersToLocal);

                // Verify that the range resolved from the name points to the expected cells
                Aspose.Cells.Range range = namedRange.GetRange();
                Console.WriteLine("Resolved range address: " + range.RefersTo);
                Console.WriteLine("Values in the resolved range:");
                for (int r = range.FirstRow; r <= range.FirstRow + range.RowCount - 1; r++)
                {
                    Console.WriteLine($"Cell A{r + 1}: {sheet.Cells[r, 0].Value}");
                }

                // Save the workbook (optional, just to complete lifecycle)
                string outputPath = "NamedRangeLocalizationVerification.xlsx";
                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
