using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaProjectReferencesFromTab
{
    static void Main()
    {
        // Path to the TAB‑delimited file containing reference data
        // Each line should have: Name<TAB>AbsoluteLibid<TAB>RelativeLibid
        string tabFilePath = "references.tsv";

        // Create a new workbook (will be saved as macro‑enabled later)
        Workbook workbook = new Workbook();

        // Access the VBA project of the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Ensure the TAB file exists
        if (!File.Exists(tabFilePath))
        {
            Console.WriteLine($"File not found: {tabFilePath}");
            return;
        }

        // Read all non‑empty lines from the TAB file
        string[] lines = File.ReadAllLines(tabFilePath);
        foreach (string line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
                continue; // skip empty lines

            // Split the line by TAB character
            string[] parts = line.Split('\t');
            if (parts.Length < 3)
            {
                Console.WriteLine($"Invalid line (expected 3 columns): {line}");
                continue;
            }

            string name = parts[0];
            string absoluteLibid = parts[1];
            string relativeLibid = parts[2];

            // Add the project reference to the VBA project
            vbaProject.References.AddProjectRefrernce(name, absoluteLibid, relativeLibid);
            Console.WriteLine($"Added reference: {name}");
        }

        // Save the workbook as a macro‑enabled file to retain VBA references
        string outputPath = "WorkbookWithReferences.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}