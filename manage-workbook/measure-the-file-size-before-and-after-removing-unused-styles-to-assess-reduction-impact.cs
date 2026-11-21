// Title: Calculate Excel workbook size reduction after removing unused styles with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to get file size before and after RemoveUnusedStyles | measure XLSX size reduction after cleaning unused styles using Aspose.Cells | Aspose.Cells .NET compare original and optimized workbook file sizes | determine percentage of size saved by removing unused styles in Excel with C# | programmatically check Excel file size change after style cleanup Aspose.Cells
// Tags: Aspose.Cells RemoveUnusedStyles API | measure Excel file size C# | calculate workbook size reduction | optimize XLSX size by cleaning styles | compare original and cleaned workbook sizes

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSX workbook, records its original byte size, calls Workbook.RemoveUnusedStyles(), saves the cleaned workbook, then measures the new file size and reports both absolute and percentage reduction.
class Program
{
    static void Main()
    {
        // Paths to the original and the processed workbook
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook from the original file
        Workbook workbook = new Workbook(inputPath);

        // Measure the original file size
        long originalSize = new FileInfo(inputPath).Length;
        Console.WriteLine($"Original file size: {originalSize} bytes");

        // Remove styles that are not used in the workbook
        workbook.RemoveUnusedStyles();

        // Save the modified workbook to a new file
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Measure the new file size after removing unused styles
        long newSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"New file size after removing unused styles: {newSize} bytes");

        // Calculate and display the size reduction
        long reduction = originalSize - newSize;
        double reductionPercent = originalSize > 0 ? (double)reduction / originalSize * 100 : 0;
        Console.WriteLine($"Size reduction: {reduction} bytes ({reductionPercent:F2}%)");
    }
}
