using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLSB workbook file
        string filePath = "sample.xlsb";

        // Open the workbook; the constructor auto‑detects the XLSB format
        Workbook workbook = new Workbook(filePath);

        // Demonstrate that the workbook is loaded (e.g., print worksheet count)
        Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
    }
}

// Author: Aspose.Cells .NET example – opens an XLSB workbook by file path.