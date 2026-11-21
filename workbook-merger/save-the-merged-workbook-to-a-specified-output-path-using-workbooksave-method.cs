// Title: Saving a merged workbook to a custom file path with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save merged workbook to specific folder | How to use workbook.Save to export a combined Excel file with Aspose.Cells | Aspose.Cells .NET specify output location when saving a merged workbook | Saving a merged workbook as XLSX using Aspose.Cells in C# | Example of workbook.Save with custom path for merged Excel workbook
// Tags: Aspose.Cells workbook.Save to custom path | save merged workbook as XLSX using Aspose.Cells | C# export combined Excel workbook with Aspose.Cells | Aspose.Cells .NET specify output file location | merged workbook saving example Aspose.Cells

using Aspose.Cells;
using System;

// Demonstrates how to persist a previously merged Workbook object to a chosen file system location by calling workbook.Save with the target path and SaveFormat.Xlsx in C#.
class Program
{
    static void Main()
    {
        // Assume the merged workbook has been created and populated elsewhere
        Workbook workbook = new Workbook(); // placeholder for the actual merged workbook

        // Define the output file path where the workbook will be saved
        string outputPath = @"C:\Output\MergedWorkbook.xlsx";

        // Save the workbook to the specified path in XLSX format
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
