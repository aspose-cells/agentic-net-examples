using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook that contains various defined names
                string sourcePath = "InputWithNames.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load only defined name objects
                LoadFilter loadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);
                LoadOptions loadOptions = new LoadOptions { LoadFilter = loadFilter };
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Access the collection of defined names
                NameCollection names = workbook.Worksheets.Names;

                // Determine names that do NOT contain the word "Total" (case‑insensitive)
                var namesToRemove = names
                    .Cast<Name>()
                    .Where(n => !n.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                    .Select(n => n.Text)
                    .ToList();

                // Remove the unwanted names from the collection
                foreach (string name in namesToRemove)
                {
                    names.Remove(name);
                }

                // Verify that only names containing "Total" remain
                Console.WriteLine("Retained defined names:");
                foreach (Name name in names)
                {
                    Console.WriteLine(name.Text);
                }

                // Save the workbook with the filtered names
                string outputPath = "FilteredNames.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}