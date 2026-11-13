// Title: Set workbook CultureInfo and refresh all pivot tables for multilingual labels using Aspose.Cells for .NET
// Common Searches: how to assign a specific CultureInfo to an Aspose.Cells workbook in C# | refresh pivot tables after changing workbook culture with Aspose.Cells | generate localized pivot table field names for each Excel file using Aspose.Cells | apply different locales to multiple workbooks and update pivot data in .NET | save workbooks with culture code suffix after pivot refresh in C#
// Tags: Workbook.Settings.CultureInfo Aspose.Cells | pivot table refresh Aspose.Cells | multilingual pivot labels .NET | apply locale per workbook Aspose.Cells | save workbook with culture suffix C#

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example loads each Excel workbook, sets a distinct CultureInfo via Workbook.Settings.CultureInfo, refreshes and recalculates all pivot tables so that labels reflect the selected locale, and saves the file with a name that includes the culture code, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Define the workbooks and the culture to apply to each one
        var workbooks = new[]
        {
            new { Path = "Workbook_EN.xlsx", Culture = "en-US" },
            new { Path = "Workbook_FR.xlsx", Culture = "fr-FR" },
            new { Path = "Workbook_JP.xlsx", Culture = "ja-JP" }
        };

        foreach (var wbInfo in workbooks)
        {
            // Verify that the source file exists
            if (!File.Exists(wbInfo.Path))
            {
                Console.WriteLine($"Source file not found: {wbInfo.Path}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(wbInfo.Path);

                // Apply distinct globalization settings (culture) to the workbook
                workbook.Settings.CultureInfo = new CultureInfo(wbInfo.Culture);

                // Refresh all pivot tables so that labels reflect the new culture
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        pivot.RefreshData();
                        pivot.CalculateData();
                    }
                }

                // Save the workbook with a name that indicates the applied culture
                string outputFile = Path.GetFileNameWithoutExtension(wbInfo.Path) +
                                    "_" + wbInfo.Culture + ".xlsx";

                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{wbInfo.Path}': {ex.Message}");
            }
        }
    }
}
