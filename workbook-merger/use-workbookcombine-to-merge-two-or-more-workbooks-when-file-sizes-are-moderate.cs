using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            WorkbookCombineDemo.Run();
        }
    }

    public class WorkbookCombineDemo
    {
        public static void Run()
        {
            // Create the first workbook (source) and add some data
            Workbook sourceWorkbook = new Workbook(); // using the Workbook() constructor rule
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";
            sourceSheet.Cells["A1"].PutValue("Data from Source Workbook");

            // Create the second workbook (destination) and add some data
            Workbook destWorkbook = new Workbook(); // using the Workbook() constructor rule
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "DestinationSheet";
            destSheet.Cells["B2"].PutValue("Data from Destination Workbook");

            // Combine the source workbook into the destination workbook
            // This merges all worksheets, styles, named ranges, etc.
            destWorkbook.Combine(sourceWorkbook); // using the Combine method rule

            // Save the combined workbook to disk in XLSX format
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx); // using the Save(string, SaveFormat) rule

            // Clean up resources
            sourceWorkbook.Dispose();
            destWorkbook.Dispose();

            Console.WriteLine("Workbooks combined and saved as 'CombinedWorkbook.xlsx'.");
        }
    }
}