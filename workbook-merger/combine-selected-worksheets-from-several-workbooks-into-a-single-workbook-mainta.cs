using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CombineSelectedWorksheets
    {
        public static void Run()
        {
            // Paths of source workbooks
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

            // Define which worksheets to copy from each source workbook
            // Key: source file path, Value: list of worksheet names to copy
            var sheetsToCopy = new Dictionary<string, string[]>
            {
                { "Source1.xlsx", new[] { "Data", "Summary" } },
                { "Source2.xlsx", new[] { "Report" } }
            };

            // Create the destination workbook (empty)
            Workbook destWorkbook = new Workbook();

            // Remove the default empty worksheet if it will not be used
            destWorkbook.Worksheets.Clear();

            // Iterate over each source workbook
            foreach (string sourcePath in sourceFiles)
            {
                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);

                // Get the list of worksheet names to copy for this source file
                if (!sheetsToCopy.TryGetValue(sourcePath, out string[] sheetNames))
                    continue; // No sheets specified for this file

                foreach (string sheetName in sheetNames)
                {
                    // Access the source worksheet
                    Worksheet srcSheet = srcWorkbook.Worksheets[sheetName];
                    if (srcSheet == null)
                        continue; // Skip if the worksheet does not exist

                    // Add a new blank worksheet to the destination workbook
                    int newIndex = destWorkbook.Worksheets.Add();

                    // Get the newly added worksheet
                    Worksheet destSheet = destWorkbook.Worksheets[newIndex];

                    // Copy the contents and formatting from the source worksheet
                    destSheet.Copy(srcSheet);

                    // Ensure unique sheet name in the destination workbook
                    string newName = srcSheet.Name;
                    int duplicateCount = 1;
                    while (destWorkbook.Worksheets.Any(ws => ws.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
                    {
                        newName = $"{srcSheet.Name}_{duplicateCount}";
                        duplicateCount++;
                    }
                    destSheet.Name = newName;
                }
            }

            // Save the combined workbook
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Combined workbook saved to '{outputPath}'.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CombineSelectedWorksheets.Run();
        }
    }
}