// Title: Replace smart marker '&=Name' in merged Excel cells while retaining merge settings using Aspose.Cells for .NET
// AI Prompts: Use WorkbookDesigner to substitute the '&=Name' smart marker inside merged cells of an Excel workbook and keep the original merge ranges unchanged. | Load an .xlsx file, bind a DataSet with a Name column to smart markers, process the workbook, and save it preserving merged cell structures.
// Common Searches: Aspose.Cells keep merged cells after processing smart markers in C# | How to replace smart marker text in merged Excel cells without losing merge layout | WorkbookDesigner replace '&=Name' placeholder while preserving cell merges .NET
// Tags: WorkbookDesigner smart marker replacement | preserve merged cells Aspose.Cells | replace placeholder in merged Excel cells C# | smart markers with DataSet Aspose.Cells | retain merge layout after processing workbook

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The program loads Input.xlsx, creates a DataSet containing a Name column, assigns it as the data source for smart markers, processes the workbook with WorkbookDesigner to replace the '&=Name' placeholder in merged cells, and saves the result to Output.xlsx while preserving the original merged cell layout.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains merged cells with smart markers (e.g., "&=Name").
            Workbook workbook = new Workbook(inputPath);

            // Prepare the data source for the smart markers using a DataSet (required by WorkbookDesigner).
            DataTable table = new DataTable("Data");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("John Doe");

            DataSet dataSource = new DataSet();
            dataSource.Tables.Add(table);

            // Initialize the WorkbookDesigner (smart marker processor) with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the data source for smart markers.
            designer.SetDataSource(dataSource);

            // Execute the smart marker replacement.
            designer.Process();

            // Save the updated workbook while keeping the original merge layout intact.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
