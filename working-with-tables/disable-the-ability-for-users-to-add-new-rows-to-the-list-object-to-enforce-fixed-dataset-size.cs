// Title: How to make an Aspose.Cells ListObject (Excel table) read‑only for row insertion using C#
// Common Searches: Aspose.Cells C# prevent users from adding rows to an Excel table | disable insert‑row UI for a ListObject with Aspose.Cells .NET | create a fixed‑size table in Excel using Aspose.Cells C# | hide the insert row placeholder in an Aspose.Cells table programmatically | make Aspose.Cells ListObject read‑only for row addition
// Tags: Aspose.Cells ListObject read only rows | C# disable insert row in Excel table | Aspose.Cells fixed size table | prevent row addition Aspose.Cells | Excel table read only with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads a workbook, accesses the first ListObject, and explains that the current Aspose.Cells API does not provide a direct property to hide the Insert Row UI; it notes where such a setting would be applied if available, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Save(inputPath);
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one ListObject (table)
            if (sheet.ListObjects.Count > 0)
            {
                // Get the first ListObject
                var table = sheet.ListObjects[0];

                // Aspose.Cells does not expose a direct property to hide the Insert Row UI.
                // If needed, you can remove the insert row placeholder by clearing the
                // table's style or by handling it in the UI layer. This placeholder is
                // not added when the table is created programmatically.
                // Example placeholder for future implementation:
                // table.ShowInsertRow = false; // Not available in current API version
            }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
