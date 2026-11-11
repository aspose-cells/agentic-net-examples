// Title: Copy a specific cell range to a new workbook and protect the workbook structure with a password using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy range A1:C10 to a new workbook and set structure protection password | how to protect workbook structure after copying cells with Aspose.Cells .NET | C# Aspose.Cells example for copying a range to another workbook and applying password protection | save selected range to a new Excel file and lock workbook structure using Aspose.Cells | Aspose.Cells protect workbook structure programmatically after range copy
// Tags: copy cell range into separate workbook Aspose.Cells C# | protect workbook structure with password Aspose.Cells API | Aspose.Cells Workbook.Protect structure operation | range copy example Aspose.Cells .NET | create new workbook from existing range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads (or creates) a source workbook, defines the A1:C10 range on its first worksheet, creates a new workbook, copies the range to the same address in the new workbook, applies structure protection with a password, and saves the result as 'copied_protected.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string targetPath = "copied_protected.xlsx";

            // Ensure the source file exists; if not, create a simple workbook as placeholder
            Workbook sourceWorkbook;
            if (File.Exists(sourcePath))
            {
                sourceWorkbook = new Workbook(sourcePath);
            }
            else
            {
                sourceWorkbook = new Workbook();
                Worksheet ws = sourceWorkbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample");
                ws.Cells["B2"].PutValue(123);
                ws.Cells["C3"].PutValue(DateTime.Now);
                sourceWorkbook.Save(sourcePath);
            }

            // Define the range to copy (A1:C10 on the first worksheet)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C10");

            // Create a new workbook (contains a default worksheet)
            Workbook newWorkbook = new Workbook();

            // Get the default worksheet of the new workbook
            Worksheet targetSheet = newWorkbook.Worksheets[0];

            // Copy the source range to the same address in the new workbook
            targetSheet.Cells.CreateRange("A1").Copy(sourceRange);

            // Protect the workbook structure with a password
            newWorkbook.Protect(ProtectionType.Structure, "MySecretPassword");

            // Save the new workbook
            newWorkbook.Save(targetPath);
            Console.WriteLine($"Workbook saved successfully to '{targetPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
