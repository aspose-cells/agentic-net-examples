// Title: Logging .NET memory consumption before and after converting an Excel workbook to TIFF using Aspose.Cells (C#)
// AI Prompts: Write a C# console program that uses GC.GetTotalMemory to record memory usage before loading a workbook, after loading, and after saving it as a TIFF with Aspose.Cells. | Show how to measure and output the managed heap size at each stage of an Excel‑to‑TIFF conversion using Aspose.Cells SaveFormat.Tiff. | Create a snippet that logs memory consumption surrounding the Workbook.Save call for TIFF export in a .NET application.
// Common Searches: C# how to measure memory usage when converting Excel to TIFF with Aspose.Cells | Aspose.Cells .NET memory profiling during workbook.Save to TIFF | track .NET heap size before and after Excel to image conversion | log GC.GetTotalMemory around Aspose.Cells TIFF export in C# console app
// Tags: Aspose.Cells workbook.Save TIFF memory profiling | GC.GetTotalMemory memory logging in C# | Excel to TIFF conversion resource usage tracking | measure .NET heap size during image export | log managed memory consumption Aspose.Cells

using System;
using Aspose.Cells;

// // Demonstrates loading an Excel file with Aspose.Cells, converting it to a TIFF image, and logging managed heap memory before loading, after loading, and after saving using GC.GetTotalMemory.
class Program
{
    static void Main()
    {
        // Paths for the source Excel file and the resulting TIFF file
        string excelPath = "input.xlsx";
        string tiffPath = "output.tiff";

        // Log memory usage before loading the workbook
        long memoryBeforeLoad = GC.GetTotalMemory(true);
        Console.WriteLine($"Memory before loading workbook: {memoryBeforeLoad / 1024} KB");

        // Load the Excel workbook
        Workbook workbook = new Workbook(excelPath);

        // Log memory usage after loading the workbook
        long memoryAfterLoad = GC.GetTotalMemory(true);
        Console.WriteLine($"Memory after loading workbook: {memoryAfterLoad / 1024} KB");

        // Convert the workbook to TIFF format and save
        workbook.Save(tiffPath, SaveFormat.Tiff);

        // Log memory usage after TIFF conversion
        long memoryAfterSave = GC.GetTotalMemory(true);
        Console.WriteLine($"Memory after TIFF conversion: {memoryAfterSave / 1024} KB");
    }
}
