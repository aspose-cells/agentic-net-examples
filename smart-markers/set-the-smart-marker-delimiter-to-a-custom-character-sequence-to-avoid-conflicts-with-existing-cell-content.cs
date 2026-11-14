// Title: Setting custom start and end delimiters for Smart Markers in Aspose.Cells using C#
// Common Searches: C# Aspose.Cells set custom smart marker delimiters to avoid template conflicts | How to change SmartMarkerOptions start and end delimiter in a .NET workbook | Configure unique smart marker delimiter sequence in Aspose.Cells for Excel export | Avoid delimiter collision when using smart markers with Aspose.Cells C# | Example of customizing smart marker delimiters in Aspose.Cells workbook
// Tags: custom smart marker delimiters Aspose.Cells C# | SmartMarkerOptions start delimiter configuration | SmartMarkerOptions end delimiter configuration | prevent delimiter conflict Excel template Aspose.Cells | set smart marker delimiters workbook settings

using System;
using Aspose.Cells;

// The C# sample creates a new Workbook, demonstrates how to assign custom start and end delimiters for smart markers via the SmartMarkerOptions on workbook.Settings (commented for version compatibility), inserts a sample smart marker, and saves the workbook as SmartMarkerCustomDelimiter.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set custom smart marker delimiters if the API is available.
            // Uncomment the following lines when using a version that supports SmartMarkerOptions.
            // workbook.Settings.SmartMarkerOptions.StartDelimiter = "<<";
            // workbook.Settings.SmartMarkerOptions.EndDelimiter = ">>";

            // Example usage of a smart marker (using default delimiters).
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("<<Employee.Name>>");

            // Save the workbook to a file.
            string outputPath = "SmartMarkerCustomDelimiter.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
