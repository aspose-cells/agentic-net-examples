// Title: Freeze the first row and column in an Aspose.Cells workbook and retrieve the XLSX as a MemoryStream (C#)
// AI Prompts: Write a C# method that creates a new Aspose.Cells Workbook, applies FreezePanes to lock the top row and leftmost column, and returns the workbook saved as an XLSX MemoryStream. | Demonstrate how to combine Aspose.Cells FreezePanes with Workbook.Save to output a frozen worksheet directly to a MemoryStream without writing a file.
// Common Searches: how to freeze top row and left column with Aspose.Cells and get a MemoryStream in C# | Aspose.Cells C# save frozen worksheet to stream instead of file | example of using FreezePanes and SaveFormat.Xlsx to return a MemoryStream
// Tags: Aspose.Cells FreezePanes row column lock | save workbook to XLSX MemoryStream C# | export frozen worksheet to stream Aspose.Cells | C# create workbook with frozen panes using Aspose | memory stream output for Aspose.Cells workbook

using Aspose.Cells;
using System;
using System.IO;

// Creates a new Workbook, freezes the first row and column (B2 as the top‑left scrollable cell), saves it in XLSX format to a MemoryStream, and returns the stream.
public class FreezeAndSaveExample
{
    public MemoryStream CreateWorkbookWithFreeze()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze first row and first column (cell B2 is the top‑left of the scrollable area)
            sheet.FreezePanes(1, 1, 0, 0);

            // Save the workbook to a memory stream in XLSX format
            MemoryStream stream = new MemoryStream();
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0;
            return stream;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating workbook: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            FreezeAndSaveExample example = new FreezeAndSaveExample();
            using (MemoryStream ms = example.CreateWorkbookWithFreeze())
            {
                // Write the stream to a file for verification
                string outputPath = "FrozenWorkbook.xlsx";
                File.WriteAllBytes(outputPath, ms.ToArray());
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
