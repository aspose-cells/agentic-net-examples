// Title: Disable editing of PivotTables while allowing refresh on a protected worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect worksheet but keep PivotTable refresh enabled | How to prevent users from editing PivotTable objects in a protected Excel sheet with Aspose.Cells | Set AllowEditObject to false and AllowRefreshPivotTable to true using Aspose.Cells .NET | Protect Excel worksheet and allow pivot cache refresh in C# Aspose.Cells example | Aspose.Cells worksheet protection settings for PivotTables in .NET
// Tags: worksheet protection allowrefreshpivottable Aspose.Cells | disable pivot table editing Aspose.Cells | protect Excel sheet with password C# Aspose.Cells | reflection set AllowEditObject property C# | Aspose.Cells worksheet protection options .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, applies full worksheet protection with a password, uses reflection to set the AllowEditObject flag to false (blocking PivotTable edits) and the AllowRefreshPivotTable flag to true (permitting PivotTable data refresh), then saves the protected file to a new location.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Protect the worksheet (all protection options) – oldPassword is empty for a new protection
            worksheet.Protect(ProtectionType.All, "MyPassword", string.Empty);

            // Disallow editing of objects (including PivotTable reports) if the property exists
            // Some older Aspose.Cells versions may not expose this property; use reflection to set it safely
            var protectionObj = worksheet.Protection;
            var allowEditObjectProp = protectionObj.GetType().GetProperty("AllowEditObject");
            if (allowEditObjectProp != null && allowEditObjectProp.CanWrite)
            {
                allowEditObjectProp.SetValue(protectionObj, false);
            }

            // Enable refreshing of PivotTables if the property exists
            var allowRefreshProp = protectionObj.GetType().GetProperty("AllowRefreshPivotTable");
            if (allowRefreshProp != null && allowRefreshProp.CanWrite)
            {
                allowRefreshProp.SetValue(protectionObj, true);
            }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
