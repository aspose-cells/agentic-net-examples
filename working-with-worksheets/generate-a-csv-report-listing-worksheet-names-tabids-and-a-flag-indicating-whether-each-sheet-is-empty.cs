// Title: C# – Generate CSV Report of Worksheet Names, TabIds, and Empty Status with Aspose.Cells
// Description: Loads an Excel workbook using Aspose.Cells for .NET, iterates all worksheets, flags a sheet as empty when MaxDataRow and MaxDataColumn are zero, and writes a CSV file containing the worksheet name, its TabId, and a Boolean IsEmpty column.
// Keywords: Aspose.Cells CSV export | list worksheets Aspose.Cells | worksheet TabId C# | detect empty worksheet Aspose.Cells | C# Excel workbook report | Aspose.Cells MaxDataRow | Aspose.Cells MaxDataColumn
// Common Searches: how to export worksheet list to CSV with Aspose.Cells | C# get worksheet TabId using Aspose.Cells | check if a worksheet is empty Aspose.Cells | generate Excel worksheet summary report C# | Aspose.Cells create CSV of sheet metadata
// Developer Intent: Produce a CSV file that enumerates each worksheet’s name, TabId, and whether it contains any data.
// Use Cases: Audit a workbook to locate and document empty sheets before cleanup. | Create a concise summary for stakeholders showing sheet identifiers and data presence. | Automate preprocessing that flags empty worksheets for removal or further handling.
// AI Prompts: Write C# code with Aspose.Cells that outputs a CSV of worksheet names, TabIds, and an IsEmpty flag based on MaxDataRow and MaxDataColumn. | Extend the sample to also include the total number of used cells for each worksheet in the CSV report. | Explain the role of MaxDataRow and MaxDataColumn in determining worksheet emptiness when using Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an Excel workbook using Aspose.Cells for .NET, iterates all worksheets, flags a sheet as empty when MaxDataRow and MaxDataColumn are zero, and writes a CSV file containing the worksheet name, its TabId, and a Boolean IsEmpty column.
class WorksheetReportGenerator
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // using the provided load constructor

        // Prepare CSV content
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("WorksheetName,TabId,IsEmpty");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine if the sheet is empty.
            // A sheet is considered empty when it has no data rows and no data columns.
            bool isEmpty = sheet.Cells.MaxDataRow == 0 && sheet.Cells.MaxDataColumn == 0;

            // Append a line with the worksheet name, TabId, and emptiness flag
            csvBuilder.AppendLine($"{sheet.Name},{sheet.TabId},{isEmpty}");
        }

        // Write the CSV report to disk
        string outputPath = "WorksheetReport.csv";
        File.WriteAllText(outputPath, csvBuilder.ToString());

        Console.WriteLine($"CSV report generated at: {Path.GetFullPath(outputPath)}");
    }
}
