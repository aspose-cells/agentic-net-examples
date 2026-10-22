// Title: Copy cell formatting between two ranges in an in‑memory workbook using Aspose.Cells for .NET
// Common Searches: how to copy cell formatting from one range to another using Aspose.Cells without saving the workbook | load an Excel workbook from a byte array and copy style between ranges in C# | Aspose.Cells Range.CopyStyle example with MemoryStream | copy style between A1:B2 and C3:D4 in a workbook loaded from memory | C# Aspose.Cells copy range style without writing file to disk
// Tags: Aspose.Cells Range.CopyStyle in-memory workbook | copy cell style between ranges C# | load workbook from byte array Aspose.Cells | memory stream Excel workbook Aspose.Cells | transfer formatting without saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook from a byte array wrapped in a MemoryStream, accesses the first worksheet, creates source (A1:B2) and destination (C3:D4) ranges, and uses Range.CopyStyle to transfer the source range's formatting to the destination range—all without writing the file to disk.
public class StyleCopyExample
{
    public void CopyStyleBetweenRanges(byte[] workbookData)
    {
        try
        {
            // Load the workbook from a memory stream
            using (MemoryStream stream = new MemoryStream(workbookData))
            {
                Workbook workbook = new Workbook(stream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define source and destination ranges
                AsposeRange sourceRange = sheet.Cells.CreateRange("A1:B2");
                AsposeRange destinationRange = sheet.Cells.CreateRange("C3:D4");

                // Copy style from source to destination
                destinationRange.CopyStyle(sourceRange);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during style copy: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string filePath = "sample.xlsx";
            byte[] workbookData;

            if (File.Exists(filePath))
            {
                // Load existing workbook file
                workbookData = File.ReadAllBytes(filePath);
            }
            else
            {
                // Create a new workbook and obtain its byte array
                Workbook wb = new Workbook();
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Save(ms, SaveFormat.Xlsx);
                    workbookData = ms.ToArray();
                }
            }

            // Execute the style copy example
            StyleCopyExample example = new StyleCopyExample();
            example.CopyStyleBetweenRanges(workbookData);

            Console.WriteLine("Style copy operation completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled error: {ex.Message}");
        }
    }
}
