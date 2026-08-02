// Title: Save a Merged Workbook to a Specified Path with Aspose.Cells for .NET (C#)
// Description: Creates a source and a destination workbook, merges them using Workbook.Combine, then persists the combined workbook to a user‑defined file location with Workbook.Save and SaveFormat.Xlsx.
// Keywords: Aspose.Cells C# save merged workbook | Workbook.Combine Aspose.Cells | Workbook.Save SaveFormat.Xlsx | merge Excel files .NET | export combined workbook path | Aspose.Cells file format Xlsx | C# Excel merge and save
// Common Searches: How to save a combined workbook using Aspose.Cells C# | Aspose.Cells merge two workbooks and save as XLSX | C# code to combine Excel files and specify output folder with Aspose.Cells | Save merged Excel file to custom directory Aspose.Cells .NET | Combine workbooks and export to Xlsx with Aspose.Cells
// Developer Intent: Persist the workbook produced by Workbook.Combine to a chosen file location.
// Use Cases: Consolidate monthly reports and summary sheets into a single XLSX file for distribution. | Merge a template workbook with user‑generated data and store the result in a shared network folder. | Automate the combination of multiple Excel files into one archiveable workbook on a scheduled server task.
// AI Prompts: Write C# code that uses Aspose.Cells to combine several workbooks and save the merged result to a specific directory with error handling. | Show how to merge two workbooks, set a custom output path, and save the combined file as .xlsb using Aspose.Cells for .NET. | Explain step‑by‑step how Workbook.Combine followed by Workbook.Save creates and stores a merged Excel file.

using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndSaveDemo
{
    // Creates a source and a destination workbook, merges them using Workbook.Combine, then persists the combined workbook to a user‑defined file location with Workbook.Save and SaveFormat.Xlsx.
    class Program
    {
        static void Main()
        {
            // Create the source workbook and add sample data
            Workbook sourceWorkbook = new Workbook();
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

            // Create the destination workbook (XLSX format) and add sample data
            Workbook destinationWorkbook = new Workbook(FileFormatType.Xlsx);
            destinationWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Merge the source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook);

            // Define the output file path
            string outputPath = "CombinedWorkbook.xlsx";

            // Save the merged workbook to the specified path using the Save method with format
            destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Merged workbook saved successfully to '{outputPath}'.");
        }
    }
}
