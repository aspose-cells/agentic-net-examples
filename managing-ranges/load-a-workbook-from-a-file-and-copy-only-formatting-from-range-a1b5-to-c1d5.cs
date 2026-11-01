// Title: Copy only cell formatting from range A1:B5 to C1:D5 in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells copy only styles between two ranges in C# | C# copy cell formatting from one range to another with Aspose.Cells | How to transfer Excel formatting without values using Aspose.Cells .NET | Copy style of range A1:B5 to C1:D5 with Aspose.Cells API | Preserve cell values while copying formatting in Excel via Aspose.Cells C#
// Tags: Aspose.Cells CopyStyle method | copy formatting between ranges Aspose.Cells | Excel range style copy C# | preserve cell values while copying styles Aspose.Cells | load workbook and copy range formatting .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates one if missing), defines source (A1:B5) and destination (C1:D5) ranges, copies only the formatting using the CopyStyle method, and saves the modified workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if it exists; otherwise create a new one
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Optional: add sample data to avoid empty ranges
                ws.Cells["A1"].PutValue("Sample1");
                ws.Cells["B5"].PutValue("Sample2");
                workbook.Save(inputPath);
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("A1", "B5");
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange("C1", "D5");

            // Copy only formatting (styles) from source to destination
            sourceRange.CopyStyle(destRange);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
