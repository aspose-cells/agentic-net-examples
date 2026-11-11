// Title: Update an external named range to point to a new workbook file path using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change external named range reference after moving source workbook | How to set RefersTo property for an external named range to a new file location in Aspose.Cells | Aspose.Cells update external range path programmatically in .NET | Modify external named range sheet and address with new workbook path using Aspose.Cells C# | Relocate external data source for a named range in an Excel file with Aspose.Cells
// Tags: Aspose.Cells external named range update | C# RefersTo property external workbook | change external reference path Aspose.Cells | update named range external file location .NET | modify external range address Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, retrieves an external named range, builds a new reference string that points to a different external workbook, sheet, and cell range, assigns this string to the named range's RefersTo property, and saves the workbook.
class UpdateExternalReference
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains the external named range
            string workbookPath = @"C:\Data\MyWorkbook.xlsx";

            // Verify that the workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // New external file path that the named range should point to
            string newExternalFilePath = @"C:\Data\NewFolder\ExternalData.xlsx";

            // Name of the external named range to update
            string externalRangeName = "MyExternalRange";

            // Load the workbook
            Workbook wb = new Workbook(workbookPath);

            // Find the named range by name
            Name namedRange = wb.Worksheets.Names[externalRangeName];

            if (namedRange != null)
            {
                // Build the new reference string.
                // Format: '[C:\Data\NewFolder\ExternalData.xlsx]Sheet1'!$A$1:$B$10
                string newSheetName = "Sheet1";               // target sheet in the external workbook
                string newAddress = "$A$1:$B$10";             // target address in the external workbook
                string newRef = $"'[${newExternalFilePath}]{newSheetName}'!{newAddress}";

                // Update the reference of the named range
                namedRange.RefersTo = newRef;

                Console.WriteLine($"Named range '{externalRangeName}' updated to reference: {newRef}");
            }
            else
            {
                Console.WriteLine($"Named range '{externalRangeName}' not found.");
            }

            // Save the workbook
            wb.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
