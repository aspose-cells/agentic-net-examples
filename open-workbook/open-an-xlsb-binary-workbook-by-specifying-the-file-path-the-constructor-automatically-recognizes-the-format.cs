// Title: Load an XLSB binary workbook from a file path with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load XLSB workbook from local path | how to open a binary .xlsb file using Aspose.Cells Workbook constructor | auto detect Excel file format when loading workbook in .NET | read first worksheet name from XLSB using Aspose.Cells C#
// Tags: load xlsb workbook Aspose.Cells | Workbook constructor auto format detection .NET | open binary Excel file C# Aspose | read worksheet name from XLSB Aspose.Cells

using Aspose.Cells;
using System;

// Demonstrates loading an XLSB binary workbook by passing its file path to the Aspose.Cells Workbook constructor, which automatically detects the format, and then prints the name of the first worksheet.
class Program
{
    static void Main()
    {
        // Path to the XLSB binary workbook
        string filePath = @"C:\Path\To\YourWorkbook.xlsb";

        // Load the workbook; the constructor automatically detects the XLSB format
        Workbook workbook = new Workbook(filePath);

        // Example usage: output the name of the first worksheet
        Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
    }
}
