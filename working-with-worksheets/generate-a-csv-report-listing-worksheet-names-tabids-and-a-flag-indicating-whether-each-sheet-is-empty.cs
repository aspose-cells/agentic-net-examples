using Aspose.Cells;
using System;
using System.IO;
using System.Text;

class WorksheetReport
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // uses the provided load rule

        // Prepare CSV header
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("WorksheetName,TabId,IsEmpty");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            string name = sheet.Name;
            int tabId = sheet.TabId;

            // Determine if the sheet is empty.
            // Check a reasonable number of columns for any data.
            bool isEmpty = true;
            for (int col = 0; col < 100; col++)
            {
                // GetLastDataRow returns -1 when the column has no data.
                if (sheet.Cells.GetLastDataRow(col) >= 0)
                {
                    isEmpty = false;
                    break;
                }
            }

            // Append the information as a CSV line
            csvBuilder.AppendLine($"{name},{tabId},{isEmpty}");
        }

        // Save the CSV report (free‑form code, no specific rule for CSV saving)
        string outputPath = "WorksheetReport.csv";
        File.WriteAllText(outputPath, csvBuilder.ToString());

        Console.WriteLine($"CSV report generated at: {outputPath}");
    }
}