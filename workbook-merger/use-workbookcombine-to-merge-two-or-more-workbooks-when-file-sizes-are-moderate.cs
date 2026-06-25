using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Destination workbook – will receive the combined content
                Workbook destWorkbook = new Workbook();

                // Paths of workbooks to be merged (moderate size files)
                string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

                // Load each source workbook and combine it into the destination workbook
                foreach (string filePath in sourceFiles)
                {
                    // Verify that the source file exists before attempting to load it
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Warning: Source file not found – skipping: {filePath}");
                        continue;
                    }

                    // Load source workbook from file
                    Workbook sourceWorkbook = new Workbook(filePath);

                    // Combine the source workbook into the destination workbook
                    destWorkbook.Combine(sourceWorkbook);
                }

                // Save the combined workbook to disk in XLSX format
                destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Combined workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}