// Title: Load SXC Workbook, Rename Active Sheet, and Export to CSV with Aspose.Cells for .NET
// Description: Demonstrates how to open an OpenDocument Spreadsheet (SXC) using Aspose.Cells, rename the currently active worksheet, and save only that sheet as a CSV file in C#.
// Keywords: Aspose.Cells C# | load SXC workbook | rename active worksheet | export to CSV | SaveFormat.Csv | .NET spreadsheet conversion | OpenDocument to CSV
// Common Searches: Aspose.Cells rename active sheet after loading SXC | Convert SXC file to CSV with Aspose.Cells .NET | C# code to change worksheet name and save as CSV | How to export only the active worksheet to CSV using Aspose.Cells
// Developer Intent: Rename the active worksheet of an SXC workbook and save it as a CSV file.
// Use Cases: Standardize sheet names before converting OpenDocument spreadsheets to CSV for downstream processing. | Automate batch conversion of SXC files to CSV with custom worksheet titles. | Integrate CSV export of a specific sheet into data pipelines that require named sheets.
// AI Prompts: Generate C# code that loads an SXC file with Aspose.Cells, renames the active worksheet, and saves it as CSV. | Explain step‑by‑step how to export only the active sheet of a workbook to CSV using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to open an OpenDocument Spreadsheet (SXC) using Aspose.Cells, rename the currently active worksheet, and save only that sheet as a CSV file in C#.
class Program
{
    static void Main()
    {
        // Path to the source SXC workbook
        string sourcePath = "input.sxc";

        // Load the workbook from the SXC file
        Workbook workbook = new Workbook(sourcePath);

        // Get the currently active worksheet and rename it
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];
        activeSheet.Name = "RenamedSheet";

        // Define the output CSV file path
        string csvPath = "output.csv";

        // Save the workbook (the active sheet) as a CSV file
        workbook.Save(csvPath, SaveFormat.Csv);

        Console.WriteLine("Workbook loaded, worksheet renamed, and saved as CSV.");
    }
}
