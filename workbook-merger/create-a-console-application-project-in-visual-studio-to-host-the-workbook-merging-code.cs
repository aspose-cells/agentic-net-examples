// Title: Build a C# console application that merges two Excel workbooks using Aspose.Cells Combine and saves the result as an XLSX file
// AI Prompts: Generate a C# console program that creates a source workbook, a destination workbook, adds sample cells, calls Workbook.Combine to merge them, and writes the merged file to a specified path in XLSX format. | Create a Visual Studio .NET console project that demonstrates Aspose.Cells workbook merging, including initializing workbooks, populating data, invoking the Combine method, and saving the combined workbook.
// Common Searches: how to use Aspose.Cells Workbook.Combine in a .NET console app | C# example for merging two Excel files and exporting as XLSX with Aspose.Cells | console application code sample for combining workbooks using Aspose.Cells | save merged Excel workbook to disk using Aspose.Cells SaveFormat.Xlsx in C#
// Tags: Aspose.Cells Workbook.Combine .NET example | C# merge Excel workbooks with Aspose.Cells | save combined workbook as XLSX using Aspose.Cells | console app workbook merging Aspose.Cells | Aspose.Cells combine method usage in C#

using System;
using Aspose.Cells;

namespace WorkbookMergingConsoleApp
{
    // Console application that demonstrates merging two workbooks using Aspose.Cells
    // The sample creates two Workbook objects, writes sample data to each, merges the source into the destination with Workbook.Combine, and saves the resulting workbook as CombinedWorkbook.xlsx using SaveFormat.Xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook(); // using Workbook() constructor
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

            // Create the destination workbook (empty workbook) and add some data
            Workbook destWorkbook = new Workbook(); // using Workbook() constructor
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Merge the source workbook into the destination workbook
            // Using the Combine method as defined in the documentation
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook to disk
            // Using the Save(string, SaveFormat) overload
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged successfully. Output saved to: {outputPath}");
        }
    }
}
