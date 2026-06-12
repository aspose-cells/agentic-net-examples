using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ResetWorksheetTabIds
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets.Add("SheetA");
                workbook.Worksheets.Add("SheetB");
                workbook.Worksheets.Add("SheetC");

                // Reset TabId for each worksheet (1‑based)
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    sheet.TabId = i + 1;
                }

                // Save the workbook
                string outputPath = "ResetTabIds.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Load the saved workbook only if it exists
                if (File.Exists(outputPath))
                {
                    Workbook loaded = new Workbook(outputPath);
                    Console.WriteLine("Worksheet TabIds after reset:");
                    foreach (Worksheet ws in loaded.Worksheets)
                    {
                        Console.WriteLine($"{ws.Name}: TabId = {ws.TabId}");
                    }
                }
                else
                {
                    Console.WriteLine($"File not found: {outputPath}");
                }
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
            ResetWorksheetTabIds.Run();
        }
    }
}