using System;
using System.IO;
using Aspose.Cells;
using System.Linq;

class WorkbookPropertiesAggregator
{
    static void Main()
    {
        // List of source workbook file paths to aggregate.
        // Replace these paths with actual file locations.
        string[] sourceFiles = new string[]
        {
            @"C:\Data\Workbook1.xlsx",
            @"C:\Data\Workbook2.xlsx",
            @"C:\Data\Workbook3.xlsx"
        };

        // Define the built‑in property names we want to include in the report.
        string[] propertyNames = new string[]
        {
            "Author",
            "Title",
            "Subject",
            "Keywords",
            "Comments",
            "LastSavedBy",
            "RevisionNumber",
            "CreatedTime",
            "LastPrinted",
            "LastSavedTime",
            "Category",
            "Manager",
            "Company"
        };

        // Create a new workbook that will hold the summary report.
        using (Workbook summaryWorkbook = new Workbook())
        {
            Worksheet sheet = summaryWorkbook.Worksheets[0];
            sheet.Name = "Properties Summary";

            // Write header row.
            sheet.Cells[0, 0].PutValue("File Name");
            for (int i = 0; i < propertyNames.Length; i++)
                sheet.Cells[0, i + 1].PutValue(propertyNames[i]);

            int rowIndex = 0;
            foreach (string filePath in sourceFiles)
            {
                // Skip missing files.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found and will be skipped: {filePath}");
                    continue;
                }

                using (Workbook srcWorkbook = new Workbook(filePath))
                {
                    sheet.Cells[rowIndex + 1, 0].PutValue(Path.GetFileName(filePath));

                    foreach (int colIndex in Enumerable.Range(0, propertyNames.Length))
                    {
                        string propName = propertyNames[colIndex];
                        var prop = srcWorkbook.BuiltInDocumentProperties[propName];
                        object value = prop?.Value;
                        sheet.Cells[rowIndex + 1, colIndex + 1].PutValue(value ?? string.Empty);
                    }
                }

                rowIndex++;
            }

            // Save the consolidated summary workbook.
            string outputPath = @"C:\Data\WorkbookPropertiesSummary.xlsx";
            summaryWorkbook.Save(outputPath);
            Console.WriteLine($"Summary report generated at: {outputPath}");
        }
    }
}