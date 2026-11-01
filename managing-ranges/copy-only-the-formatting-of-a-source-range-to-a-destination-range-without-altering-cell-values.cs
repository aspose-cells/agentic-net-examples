// Title: Copy only cell formatting from source range A1:C3 to destination range E1:G3 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy only formatting from a source range to a destination range | How to transfer cell styles without changing values using Aspose.Cells .NET | Copy Excel range formatting to another area while keeping original data in C# | Range.CopyStyle method example for non‑adjacent ranges in Aspose.Cells | Preserve cell values when copying styles with Aspose.Cells for .NET
// Tags: Aspose.Cells Range.CopyStyle method example | Excel range formatting transfer C# | copy cell style without data modification | nonadjacent range style copy .NET | preserve data during style copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample creates or loads an Excel workbook, defines a source range (A1:C3) and a destination range (E1:G3) on the first worksheet, and uses the Range.CopyStyle method to copy only the formatting while leaving cell values unchanged, then saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a sample workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var sampleWorkbook = new Workbook();
                var ws = sampleWorkbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["C1"].PutValue("Header3");
                ws.Cells["A2"].PutValue(1);
                ws.Cells["B2"].PutValue(2);
                ws.Cells["C2"].PutValue(3);
                ws.Cells["A3"].PutValue(4);
                ws.Cells["B3"].PutValue(5);
                ws.Cells["C3"].PutValue(6);
                sampleWorkbook.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // Get the first worksheet.
            var sheet = workbook.Worksheets[0];

            // Define the source range (A1:C3).
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1", "C3");

            // Define the destination range (E1:G3).
            AsposeRange destinationRange = sheet.Cells.CreateRange("E1", "G3");

            // Copy only the formatting from source to destination.
            sourceRange.CopyStyle(destinationRange);

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
