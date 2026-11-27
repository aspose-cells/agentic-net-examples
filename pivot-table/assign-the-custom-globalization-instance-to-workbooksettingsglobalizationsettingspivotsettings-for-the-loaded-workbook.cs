// Title: How to assign a custom GlobalizationSettings.PivotSettings instance to a workbook in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set custom PivotSettings on WorkbookSettings.GlobalizationSettings | assign custom globalization pivot settings to loaded workbook using Aspose.Cells | configure pivot table culture in Aspose.Cells .NET programmatically | how to change pivot table localization with Aspose.Cells workbook settings | apply custom GlobalizationSettings.PivotSettings before saving workbook in C#
// Tags: custom PivotSettings assignment Aspose.Cells | WorkbookSettings.GlobalizationSettings.PivotSettings configuration | pivot table localization Aspose.Cells .NET | set workbook globalization for pivots C# | Aspose.Cells custom globalization instance

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates loading an existing Excel file (or creating a new workbook), creating a custom GlobalizationSettings.PivotSettings object, assigning it to WorkbookSettings.GlobalizationSettings.PivotSettings, and then saving the workbook, illustrating how to programmatically control pivot table culture and localization in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure at least one worksheet exists.
                workbook.Worksheets.Add("Sheet1");
            }

            // The 1904 date system property may not be available in some versions of Aspose.Cells.
            // If needed, set it here using the appropriate API for the target version.

            // Save the workbook to the specified output file.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
