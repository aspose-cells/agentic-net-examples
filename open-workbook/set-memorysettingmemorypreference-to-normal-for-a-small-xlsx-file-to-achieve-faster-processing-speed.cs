// Title: C# – Set Workbook MemorySetting to Normal for Faster Small XLSX Processing with Aspose.Cells
// Description: Demonstrates how to create a workbook, apply MemorySetting.Normal to boost performance on a small XLSX file, add sample data, and save the document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells MemorySetting Normal | C# Excel performance small file | Workbook.Settings.MemorySetting | optimize Aspose.Cells speed | fast XLSX generation .NET
// Common Searches: Aspose.Cells set MemorySetting to Normal | increase speed for small Excel files Aspose | MemorySetting.Normal C# example | best memory preference for tiny workbooks Aspose.Cells
// Developer Intent: Apply the Normal memory preference to a workbook so that small Excel files are processed more quickly.
// Use Cases: Generate lightweight reports in a web service where latency must be minimal. | Create dozens of temporary spreadsheets in a batch job without excessive memory overhead. | Build on‑the‑fly Excel files for API responses, using MemorySetting.Normal to keep response time low.
// AI Prompts: Write C# code that opens an existing small XLSX file with Aspose.Cells, sets MemorySetting to Normal, and then reads its contents. | Explain how MemorySetting.Normal affects memory consumption and execution time for small versus large workbooks in Aspose.Cells. | Provide a conditional snippet that chooses MemorySetting.Normal or MemorySetting.High based on the file size before loading a workbook.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply MemorySetting.Normal to boost performance on a small XLSX file, add sample data, and save the document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set memory usage to Normal for faster processing on a small file
        workbook.Settings.MemorySetting = MemorySetting.Normal;

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Save the workbook to an XLSX file
        workbook.Save("MemorySettingNormal.xlsx");
    }
}
