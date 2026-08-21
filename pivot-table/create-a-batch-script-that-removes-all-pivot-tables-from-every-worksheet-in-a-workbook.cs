// Title: C# – Remove All Pivot Tables from an Excel Workbook Using Aspose.Cells
// Description: Load an Excel file with Aspose.Cells, call Worksheets.ClearPivottables() to delete every pivot table across all worksheets, and save the cleaned workbook. Ideal for batch cleanup of reports or templates.
// Keywords: Aspose.Cells clear pivot tables | remove all pivot tables .NET | delete pivot tables workbook | C# Aspose.Cells pivot table removal | Excel pivot table cleanup
// Common Searches: how to delete all pivot tables with Aspose.Cells | Aspose.Cells remove pivot tables from every sheet | C# batch script to clear pivot tables in Excel | clear pivot tables across workbook Aspose.Cells
// Developer Intent: Eliminate every pivot table in a workbook and save the result.
// Use Cases: Sanitize a template before distribution by stripping all pivot tables. | Prepare data‑only Excel files for downstream processing. | Automate cleanup of generated reports that contain unwanted pivot tables.
// AI Prompts: Write C# code that uses Aspose.Cells to remove all pivot tables from an Excel workbook and saves the file. | Explain the requirements and side effects of Worksheets.ClearPivottables() in Aspose.Cells. | Create a PowerShell script that runs a compiled .NET executable to batch‑process multiple Excel files, removing their pivot tables.

using System;
using Aspose.Cells;

// Load an Excel file with Aspose.Cells, call Worksheets.ClearPivottables() to delete every pivot table across all worksheets, and save the cleaned workbook. Ideal for batch cleanup of reports or templates.
class Program
{
    static void Main(string[] args)
    {
        // Input Excel file containing pivot tables
        string inputPath = "input.xlsx";

        // Output Excel file after removing all pivot tables
        string outputPath = "output_no_pivots.xlsx";

        // Load the workbook from the input file
        Workbook workbook = new Workbook(inputPath);

        // Clear all pivot tables from every worksheet in the workbook
        workbook.Worksheets.ClearPivottables();

        // Save the modified workbook to the output file
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
