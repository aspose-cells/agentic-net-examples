// Title: Load an XLSB binary workbook from a file path with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load XLSB workbook from local file path | How to open a binary Excel (.xlsb) file using Aspose.Cells Workbook constructor | Aspose.Cells automatically detect file format when creating Workbook in .NET | Read first worksheet name after loading XLSB with Aspose.Cells | Open Excel binary workbook without specifying format in C# Aspose.Cells
// Tags: load xlsb workbook Aspose.Cells | Workbook constructor format detection .NET | read worksheet name C# Aspose.Cells | open binary Excel file using Aspose.Cells | Aspose.Cells file path initialization | C# load binary Excel workbook

using System;
using Aspose.Cells;

// The example shows how to specify a file path to an XLSB workbook, instantiate a Workbook object that auto‑detects the binary format, and then retrieve the name of the first worksheet using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Specify the full path to the XLSB workbook.
        string filePath = @"C:\Path\To\YourWorkbook.xlsb";

        // The Workbook constructor automatically detects the file format (XLSB in this case) and loads it.
        Workbook workbook = new Workbook(filePath);

        // At this point the workbook is loaded and can be manipulated.
        // Example: output the name of the first worksheet.
        Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
    }
}
