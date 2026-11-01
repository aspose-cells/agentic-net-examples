// Title: How to assign a different CultureInfo to each worksheet in an Aspose.Cells workbook using C#
// Common Searches: C# Aspose.Cells set French locale for a specific worksheet while keeping English for others | How to change workbook.Settings.CultureInfo per sheet in Aspose.Cells .NET | Apply different number and date formats per worksheet with Aspose.Cells in C# | Save an Excel file with English and French worksheets using separate CultureInfo objects in Aspose.Cells | Switch culture between worksheets when populating data with Aspose.Cells for .NET
// Tags: worksheet-specific CultureInfo Aspose.Cells | multilingual Excel workbook C# Aspose.Cells | set culture per sheet .NET | locale-aware number formatting Aspose.Cells | date formatting per worksheet Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets the default culture to en‑US, adds an English worksheet and writes a numeric and date value, switches the workbook’s CultureInfo to fr‑FR before populating a French worksheet, restores en‑US for subsequent operations, and saves the file as LocalizedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and set default culture to English (US)
            Workbook workbook = new Workbook();
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // First worksheet: English (US) localization
            Worksheet wsEnglish = workbook.Worksheets[0];
            wsEnglish.Name = "English";

            // Populate English worksheet
            wsEnglish.Cells["A1"].PutValue(12345.67);          // Number
            wsEnglish.Cells["A2"].PutValue(DateTime.Now);    // Date

            // Second worksheet: French (France) localization
            Worksheet wsFrench = workbook.Worksheets.Add("French");

            // Switch culture to French for French worksheet data entry
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Populate French worksheet
            wsFrench.Cells["A1"].PutValue(12345.67);
            wsFrench.Cells["A2"].PutValue(DateTime.Now);

            // Restore culture to English for any further operations
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // Define output file path
            string outputPath = "LocalizedWorkbook.xlsx";

            // Delete existing file to avoid conflicts
            try
            {
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not delete existing file. {ex.Message}");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
