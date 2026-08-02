// Title: Configure Aspose.Cells MemorySetting.Normal for Faster Small XLSX Processing (C#)
// Description: Creates a new Workbook, sets Settings.MemorySetting to MemorySetting.Normal to speed up handling of a small Excel file, adds sample text and a date, then saves as SmallFile_NormalMemory.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | MemorySetting.Normal | C# | .NET Excel performance | small workbook speed | fast Excel generation | memory mode optimization | Workbook.Settings.MemorySetting | Excel file processing speed | low memory usage Aspose
// Common Searches: Aspose.Cells set MemorySetting to Normal | How to improve performance for small XLSX with Aspose.Cells .NET | MemorySetting.Normal vs LowMemory Aspose.Cells | Fast workbook creation C# Aspose.Cells | Optimize memory settings for Excel files Aspose
// Developer Intent: Apply the Normal memory mode to a workbook so a small XLSX file is processed more quickly.
// Use Cases: Generating lightweight reports where rapid workbook creation is critical. | Batch‑processing dozens of small Excel files in a background service with minimal latency. | Creating temporary Excel files in a web API that require fast response and modest memory consumption.
// AI Prompts: Show how to set Aspose.Cells MemorySetting.Normal in C# and explain its impact on small workbook performance. | Provide a concise example that configures MemorySetting.Normal, adds data, and saves an XLSX file. | Compare MemorySetting.Normal with LowMemory and advise when each should be used in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new Workbook, sets Settings.MemorySetting to MemorySetting.Normal to speed up handling of a small Excel file, adds sample text and a date, then saves as SmallFile_NormalMemory.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set memory usage to Normal for faster processing of a small file
        workbook.Settings.MemorySetting = MemorySetting.Normal;

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Save the workbook
        workbook.Save("SmallFile_NormalMemory.xlsx");
    }
}
