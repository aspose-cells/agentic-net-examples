// Title: Unmerge cells A1:C1 and copy them to a new worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unmerge a merged range and copy its values to another sheet | How to copy unmerged cells from one worksheet to a new worksheet using Aspose.Cells | C# code to unmerge A1:C1 and duplicate the range on a different worksheet with Aspose.Cells | Aspose.Cells example for unmerging cells and copying formatting to a new sheet in .NET | Copy merged header row after unmerging with Aspose.Cells for .NET
// Tags: unmerge cells Aspose.Cells C# | copy range between worksheets Aspose.Cells | create new worksheet Aspose.Cells | transfer cell formatting Aspose.Cells | handle merged header row Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads or creates Input.xlsx, unmerges the A1:C1 range on the first worksheet, copies its values, formulas, and formatting to a newly added worksheet named 'CopySheet', and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Load existing workbook or create a new one if the input file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";

                // Populate sample data in A1:C1
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["C1"].PutValue("Header3");

                // Merge A1:C1 to simulate the original scenario
                ws.Cells.Merge(0, 0, 1, 3);

                // Save the temporary input file (optional)
                workbook.Save(inputPath);
            }

            // Get the first worksheet (source)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Unmerge the range A1:C1 in the source sheet using Aspose.Cells.Range
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:C1");
            sourceRange.UnMerge();

            // Add a new worksheet to copy the content into (target)
            Worksheet targetSheet = workbook.Worksheets.Add("CopySheet");

            // Create corresponding range in the target sheet
            Aspose.Cells.Range targetRange = targetSheet.Cells.CreateRange("A1:C1");

            // Copy the content (values, formulas, formatting) from source to target
            targetRange.Copy(sourceRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
