// Title: Save an Excel workbook with ISO‑29500‑2008 strict compliance using Aspose.Cells for .NET and check the file size in C#
// AI Prompts: Create a new Workbook, set Settings.Compliance to OoxmlCompliance.Iso29500_2008_Strict, write a value to a cell, save it as an .xlsx file, and output the resulting file size. | Load the saved .xlsx file with the Workbook constructor and display its Settings.Compliance value to confirm strict ISO compliance.
// Common Searches: Aspose.Cells C# how to enforce ISO 29500-2008 strict mode when saving an Excel file | C# code to get file size after saving workbook with Aspose.Cells | Verify workbook compliance setting after loading a strict OOXML file using Aspose.Cells | Set OoxmlCompliance to Iso29500_2008_Strict in Aspose.Cells .NET example | Check saved Excel file size programmatically with Aspose.Cells in C#
// Tags: aspocells set workbook compliance iso29500 strict | aspocells save workbook strict ooxml | c# get saved excel file size aspocells | aspocells load workbook read compliance setting | ooxml compliance iso29500_2008_strict c#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, configures its Settings.Compliance to ISO‑29500‑2008 strict mode, writes a sample value, saves the file as StrictCompliance.xlsx, prints the saved file size, then reloads the workbook to display the compliance setting, demonstrating both strict OOXML compliance and file‑size verification.
class Program
{
    static void Main()
    {
        // Create a new workbook (uses the Workbook() constructor rule)
        Workbook workbook = new Workbook();

        // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Add a simple value so the file is not empty
        workbook.Worksheets[0].Cells["A1"].PutValue("Strict compliance demo");

        // Define the output file name
        string outputPath = "StrictCompliance.xlsx";

        // Save the workbook (uses the Save(string) rule)
        workbook.Save(outputPath);

        // Verify the saved file size
        FileInfo fileInfo = new FileInfo(outputPath);
        Console.WriteLine($"Saved file size: {fileInfo.Length} bytes");

        // Load the workbook back to confirm it can be opened (uses Workbook(string) rule)
        Workbook loadedWorkbook = new Workbook(outputPath);
        Console.WriteLine($"Loaded workbook compliance: {loadedWorkbook.Settings.Compliance}");
    }
}
