// Title: C# – Use Aspose.Cells ShowReportFilterPage to Create a Worksheet per Pivot Table Report Filter
// Description: Loads a workbook, validates the first worksheet contains a pivot table, iterates through each report‑filter (page) field, calls ShowReportFilterPage to generate a separate worksheet for every filter value, and saves the updated file. Ideal for splitting pivot data into individual sheets programmatically.
// Keywords: Aspose.Cells ShowReportFilterPage | C# pivot table report filter worksheets | Aspose.Cells generate sheets per filter | PivotTable page fields .NET | split pivot data into separate worksheets | Aspose.Cells workbook save example
// Common Searches: Aspose.Cells ShowReportFilterPage C# example | create worksheet for each pivot report filter .NET | generate separate sheets from pivot table filters | Aspose.Cells split pivot table by page field | how to export pivot filter pages using Aspose.Cells
// Developer Intent: Programmatically produce one worksheet for each report‑filter selection of a pivot table and save the workbook.
// Use Cases: Produce regional sales reports by creating a sheet for each "Region" filter. | Export product‑category data into individual worksheets for targeted analysis. | Automate department‑specific reporting from a master pivot table by generating a sheet per "Department" filter.
// AI Prompts: Show a C# code snippet that uses Aspose.Cells ShowReportFilterPage to create worksheets for all pivot table page fields. | Explain the parameters and customization options available with ShowReportFilterPage in Aspose.Cells. | Suggest robust error‑handling patterns for loading workbooks and verifying pivot tables before invoking ShowReportFilterPage.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads a workbook, validates the first worksheet contains a pivot table, iterates through each report‑filter (page) field, calls ShowReportFilterPage to generate a separate worksheet for every filter value, and saves the updated file. Ideal for splitting pivot data into individual sheets programmatically.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the first worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // For each page field (report filter) create a separate worksheet
            // that contains the data for that filter selection
            foreach (PivotField pageField in pivotTable.PageFields)
            {
                // This method creates a new worksheet with the filtered data
                pivotTable.ShowReportFilterPage(pageField);
            }

            // Save the workbook with the newly generated worksheets
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
