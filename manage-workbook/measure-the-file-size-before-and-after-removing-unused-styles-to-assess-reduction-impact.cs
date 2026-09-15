// Title: Measure Excel workbook size before and after removing unused styles with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that loads an XLSX file with Aspose.Cells, records its byte size, calls Workbook.RemoveUnusedStyles(), saves the workbook, and outputs the size difference. | Show how to compare the file size of an Excel workbook before and after cleaning up unused styles using Aspose.Cells in a .NET console application. | Demonstrate logging the reduction in bytes achieved by Workbook.RemoveUnusedStyles() for a given workbook with Aspose.Cells.
// Common Searches: how to compare Excel file size before and after RemoveUnusedStyles in C# | Aspose.Cells C# example for measuring workbook size reduction after style cleanup | C# code to get XLSX byte size, remove unused styles, and see impact | does Workbook.RemoveUnusedStyles reduce file size of large Excel workbooks | measure file size change after cleaning up Excel styles with Aspose.Cells .NET
// Tags: Aspose.Cells RemoveUnusedStyles file size impact C# | Excel workbook byte reduction after style cleanup | C# track XLSX size before after Aspose.Cells | optimize workbook unused styles Aspose.Cells | evaluate workbook size change Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' using Aspose.Cells, records its byte size, removes all unused styles with Workbook.RemoveUnusedStyles(), saves the cleaned workbook as 'output.xlsx', and prints the file sizes before and after the cleanup to illustrate the size reduction.
class Program
{
    static void Main()
    {
        // Path to the original workbook
        string inputPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Measure file size before removing unused styles
        long sizeBefore = new FileInfo(inputPath).Length;
        Console.WriteLine($"File size before removing unused styles: {sizeBefore} bytes");

        // Remove all styles that are not used in the workbook
        workbook.RemoveUnusedStyles();

        // Save the modified workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        // Measure file size after removal
        long sizeAfter = new FileInfo(outputPath).Length;
        Console.WriteLine($"File size after removing unused styles: {sizeAfter} bytes");
    }
}
