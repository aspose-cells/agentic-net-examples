// Title: Set PreserveFormatting = true for all QueryTables in every worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates over each worksheet, sets QueryTable.PreserveFormatting to true for every query table, and saves the file. | Show a step‑by‑step example of enabling formatting preservation on all data‑connection tables in a workbook and exporting it as XLSX with Aspose.Cells.
// Common Searches: Aspose.Cells C# set PreserveFormatting on all query tables in a workbook | How to enable formatting preservation for Excel QueryTable objects using Aspose.Cells | Iterate through worksheets and update QueryTable properties in .NET | Batch modify QueryTable.PreserveFormatting flag with Aspose.Cells API | Preserve cell formatting when refreshing data connections in Aspose.Cells
// Tags: Aspose.Cells QueryTable formatting preservation | C# loop through worksheets to modify QueryTables | batch set QueryTable properties Aspose.Cells | Excel data connection formatting preservation .NET | enable QueryTable PreserveFormatting flag

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an existing workbook (or creates a new one if missing), walks through every worksheet's QueryTable collection, sets each QueryTable's PreserveFormatting property to true, and saves the updated workbook as output.xlsx.
    public class SetQueryTablePreserveFormatting
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                Console.WriteLine($"Input file '{inputPath}' not found. A new workbook has been created.");
            }

            // Iterate through all worksheets and set PreserveFormatting for each QueryTable
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                QueryTableCollection queryTables = sheet.QueryTables;
                for (int i = 0; i < queryTables.Count; i++)
                {
                    QueryTable qt = queryTables[i];
                    qt.PreserveFormatting = true;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
