// Title: C# – List Worksheets with >10 Shapes and Export a Summary Report using Aspose.Cells
// Description: Loads a workbook, scans each worksheet’s ShapeCollection, gathers the names of sheets containing more than ten shapes, writes those names to a new workbook (ComplexSheetsReport.xlsx), and prints the list to the console.
// Keywords: Aspose.Cells | C# | shape count | worksheet shapes | list worksheets | generate Excel report | shape collection | filter worksheets by shapes | Excel automation | performance audit
// Common Searches: Aspose.Cells count shapes in worksheet C# | C# find worksheets with many shapes using Aspose | generate Excel report of sheets with high shape count | list worksheets exceeding shape threshold Aspose.Cells | how to audit shape count in Excel with Aspose
// Developer Intent: Identify worksheets that exceed a shape‑count threshold and create a separate Excel file summarizing those sheet names.
// Use Cases: Audit a workbook to locate sheets with extensive graphical content before redesign. | Generate a stakeholder report highlighting worksheets that may affect performance due to many shapes. | Add an automated quality‑check step in CI pipelines to flag sheets surpassing a shape limit. | Document complex worksheets for migration or refactoring projects. | Create a quick inventory of graphic‑heavy sheets for resource planning.
// AI Prompts: Write C# Aspose.Cells code that lists worksheet names where ShapeCollection.Count > 10 and saves the list to a new Excel file. | Show how to modify the snippet to include the exact shape count next to each worksheet name in the report. | Explain how to make the shape threshold and output file path configurable parameters. | Provide a version that logs the worksheet names and counts to a text file instead of the console. | Create a PowerShell script that compiles and runs the C# program, then opens the generated report.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads a workbook, scans each worksheet’s ShapeCollection, gathers the names of sheets containing more than ten shapes, writes those names to a new workbook (ComplexSheetsReport.xlsx), and prints the list to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (lifecycle: load)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Collect names of worksheets that contain more than 10 shapes
        List<string> complexSheets = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Worksheet.Shapes returns the ShapeCollection; Count gives the number of shapes
            if (sheet.Shapes.Count > 10)
            {
                complexSheets.Add(sheet.Name);
            }
        }

        // Output the result to the console
        Console.WriteLine("Worksheets with more than 10 shapes:");
        foreach (string name in complexSheets)
        {
            Console.WriteLine("- " + name);
        }

        // Create a new workbook to store a simple report (lifecycle: create)
        Workbook report = new Workbook();
        Worksheet reportSheet = report.Worksheets[0];
        reportSheet.Name = "ComplexSheets";

        // Write header
        reportSheet.Cells[0, 0].PutValue("Worksheet Name");

        // Write each worksheet name into the report sheet
        int row = 1;
        foreach (string name in complexSheets)
        {
            reportSheet.Cells[row, 0].PutValue(name);
            row++;
        }

        // Save the report workbook to disk (lifecycle: save)
        report.Save("ComplexSheetsReport.xlsx");
    }
}
