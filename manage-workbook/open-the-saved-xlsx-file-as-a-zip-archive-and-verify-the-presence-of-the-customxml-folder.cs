// Title: Check for a customXml folder in an .xlsx workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to detect custom XML parts in an .xlsx file | C# code to verify if an Excel workbook contains a customXml folder | Determine presence of customXml folder in .xlsx using Aspose.Cells API | Inspect Excel package for custom XML parts with Aspose.Cells .NET | Load workbook and check customXml folder existence in C#
// Tags: Aspose.Cells custom XML parts detection | C# check customXml folder in Excel workbook | load .xlsx workbook inspect customXml parts | Excel package custom XML verification with Aspose | customXml folder presence validation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx file with Aspose.Cells, examines the Workbook.CustomXmlParts collection, and outputs whether the customXml folder (custom XML parts) is present in the workbook.
class Program
{
    static void Main()
    {
        // Path to the .xlsx file to be inspected
        string filePath = @"C:\Temp\MyWorkbook.xlsx";

        try
        {
            // Verify that the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(filePath);

            // Determine if any custom XML parts are present
            bool hasCustomXml = workbook.CustomXmlParts != null && workbook.CustomXmlParts.Count > 0;

            // Output the verification result
            Console.WriteLine(hasCustomXml
                ? "customXml folder is present in the workbook."
                : "customXml folder is NOT present in the workbook.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
