// Title: Save a modified Excel workbook to a new XLSX file while preserving all original formatting using Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX workbook, change the value of cell A1, and save it as a new file without affecting any existing styles using Aspose.Cells in C#. | Create C# code that opens a workbook, updates a cell, and exports the workbook to XLSX while keeping every original formatting element intact with Aspose.Cells.
// Common Searches: Aspose.Cells C# save workbook while retaining cell formatting | How to keep original styles when exporting a modified Excel file with Aspose.Cells | C# example to load an XLSX, edit a cell, and save to a new file without losing formatting | Using SaveFormat.Xlsx to preserve formatting in Aspose.Cells .NET
// Tags: Aspose.Cells workbook.Save retain formatting | C# modify cell and export to XLSX with original styles | SaveFormat.Xlsx for style retention Aspose.Cells | load existing workbook keep cell styles Aspose.Cells .NET | Aspose.Cells maintain formatting during save

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // // Loads an existing XLSX (or creates a new workbook), changes cell A1 to "Modified", and saves the workbook as a new XLSX file using Aspose.Cells, ensuring all original formatting and styles are retained.
    public class SaveWorkbookDemo
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Load existing workbook if it exists; otherwise create a new one.
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Example modification: change the value of a cell.
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Modified");

                // Save the workbook to a new XLSX file, preserving formatting.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SaveWorkbookDemo.Run();
        }
    }
}
