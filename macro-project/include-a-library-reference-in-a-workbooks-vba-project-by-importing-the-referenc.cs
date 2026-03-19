using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceImport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the CSV file containing reference definitions.
            // Expected format per line: Name,AbsoluteLibid,RelativeLibid
            string csvPath = "references.csv";

            // Create a new workbook. Saving as Xlsm will ensure a VBA project is present.
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            // Read all lines from the CSV file.
            if (File.Exists(csvPath))
            {
                string[] lines = File.ReadAllLines(csvPath);

                foreach (string line in lines)
                {
                    // Skip empty lines.
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Split the line by commas. Trim each part to remove extra whitespace.
                    string[] parts = line.Split(',');

                    // Expect exactly three columns: name, absoluteLibid, relativeLibid.
                    if (parts.Length != 3)
                    {
                        Console.WriteLine($"Invalid line format (expected 3 columns): {line}");
                        continue;
                    }

                    string name = parts[0].Trim();
                    string absoluteLibid = parts[1].Trim();
                    string relativeLibid = parts[2].Trim();

                    // Add the project reference to the VBA project.
                    // The method returns the index of the added reference (unused here).
                    vbaProject.References.AddProjectRefrernce(name, absoluteLibid, relativeLibid);
                    Console.WriteLine($"Added reference: {name}");
                }
            }
            else
            {
                Console.WriteLine($"CSV file not found at path: {csvPath}");
            }

            // Save the workbook as a macro-enabled file to preserve the VBA project and its references.
            string outputPath = "WorkbookWithReferences.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}