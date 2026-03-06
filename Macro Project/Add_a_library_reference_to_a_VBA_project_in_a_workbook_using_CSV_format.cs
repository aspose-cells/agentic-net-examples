using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the CSV file containing reference data.
            // Expected format per line: Name,AbsoluteLibid,RelativeLibid
            string csvPath = "references.csv";

            // Create a new workbook (macro-enabled format will be used when saving).
            Workbook workbook = new Workbook();

            // Ensure the workbook has a VBA project (it is created automatically for .xlsm files).
            VbaProject vbaProject = workbook.VbaProject;

            // Read each line from the CSV and add a project reference.
            if (File.Exists(csvPath))
            {
                foreach (string line in File.ReadAllLines(csvPath))
                {
                    // Skip empty lines.
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Split the CSV line into parts.
                    string[] parts = line.Split(',');

                    // Validate that we have exactly three columns.
                    if (parts.Length != 3)
                    {
                        Console.WriteLine($"Invalid CSV line (expected 3 columns): {line}");
                        continue;
                    }

                    string name = parts[0].Trim();
                    string absoluteLibid = parts[1].Trim();
                    string relativeLibid = parts[2].Trim();

                    // Add the project reference using Aspose.Cells VBA API.
                    // The method returns the index of the added reference (unused here).
                    vbaProject.References.AddProjectRefrernce(name, absoluteLibid, relativeLibid);
                    Console.WriteLine($"Added reference: {name}");
                }
            }
            else
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
            }

            // Save the workbook as a macro-enabled file.
            string outputPath = "WorkbookWithVbaReference.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}