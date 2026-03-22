using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public static class CombineSelectedWorksheetsDemo
    {
        public static void CombineSelectedWorksheets(string[] sourceFiles, Dictionary<string, List<string>> sheetsToCopy, string outputFile)
        {
            // Create the destination workbook (empty workbook)
            Workbook destWorkbook = new Workbook();

            // Iterate over each source file
            foreach (string srcPath in sourceFiles)
            {
                // Load the source workbook
                Workbook srcWorkbook = new Workbook(srcPath);

                // Get the list of worksheet names to copy for this source file
                if (!sheetsToCopy.TryGetValue(srcPath, out List<string> sheetNames) || sheetNames == null)
                {
                    // If no specific list is provided, skip this file
                    continue;
                }

                // Copy each requested worksheet into the destination workbook
                foreach (string sheetName in sheetNames)
                {
                    // Verify that the worksheet exists in the source workbook
                    Worksheet srcSheet = srcWorkbook.Worksheets[sheetName];
                    if (srcSheet == null)
                    {
                        Console.WriteLine($"Worksheet '{sheetName}' not found in '{srcPath}'. Skipping.");
                        continue;
                    }

                    // Add a copy of the worksheet to the destination workbook using the sheet name overload
                    int newIndex = destWorkbook.Worksheets.AddCopy(srcSheet.Name);

                    // Optionally rename the copied sheet to avoid name collisions
                    Worksheet copiedSheet = destWorkbook.Worksheets[newIndex];
                    copiedSheet.Name = $"{srcWorkbook.Worksheets[0].Name}_{sheetName}";
                }
            }

            // Save the combined workbook to the specified output path
            destWorkbook.Save(outputFile, SaveFormat.Xlsx);
        }

        public static void Run()
        {
            // Define source files
            string[] files = { "Source1.xlsx", "Source2.xlsx" };

            // Define which worksheets to copy from each source file
            var sheetsMap = new Dictionary<string, List<string>>
            {
                { "Source1.xlsx", new List<string> { "Data", "Summary" } },
                { "Source2.xlsx", new List<string> { "Report" } }
            };

            // Output file path
            string resultPath = "CombinedWorkbook.xlsx";

            // Perform the combination
            CombineSelectedWorksheets(files, sheetsMap, resultPath);

            Console.WriteLine($"Combined workbook saved to '{resultPath}'.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CombineSelectedWorksheetsDemo.Run();
        }
    }
}