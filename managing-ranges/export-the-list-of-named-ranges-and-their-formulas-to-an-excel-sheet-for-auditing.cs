using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeAuditor
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "AuditedWorkbook.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet to hold the audit information
                Worksheet auditSheet = workbook.Worksheets.Add("Audit");
                Cells auditCells = auditSheet.Cells;

                // Write header titles
                auditCells["A1"].PutValue("Named Range");
                auditCells["B1"].PutValue("Refers To");

                // Retrieve all defined names (named ranges) in the workbook
                NameCollection names = workbook.Worksheets.Names;

                // Iterate through each name and write its details to the audit sheet
                int row = 1; // zero‑based index; start after header
                foreach (Name name in names)
                {
                    auditCells[row, 0].PutValue(name.Text);      // Column A: name text
                    auditCells[row, 1].PutValue(name.RefersTo); // Column B: formula the name refers to
                    row++;
                }

                // Save the workbook with the audit sheet added
                workbook.Save(outputPath);
                Console.WriteLine($"Audit completed. Saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}