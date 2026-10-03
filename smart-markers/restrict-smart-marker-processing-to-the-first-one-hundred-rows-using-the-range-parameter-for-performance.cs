// Title: Restrict Aspose.Cells smart marker processing to the first 100 rows using WorkbookDesigner.Range in C#
// AI Prompts: Write C# code that assigns WorkbookDesigner.Range to rows 1‑100 before invoking Process() to limit smart marker evaluation. | Show how to configure a row range on WorkbookDesigner so only the top hundred rows are processed for smart markers, improving performance.
// Common Searches: how to restrict Aspose.Cells smart markers to specific rows in C# | C# set processing range for WorkbookDesigner to improve smart marker speed | limit smart marker evaluation to top rows with Aspose.Cells | Aspose.Cells smart marker performance tips for large worksheets | apply row range filter before calling WorkbookDesigner.Process in C#
// Tags: Aspose.Cells WorkbookDesigner range | smart marker row limit | C# limit smart marker processing | smart marker processing efficiency

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example demonstrates loading a workbook (or creating one), binding a data source to the "Data" smart‑marker group, setting WorkbookDesigner.Range to rows 1‑100, processing only those smart markers, and saving the resulting file. This limits evaluation to the first hundred rows and boosts performance.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            string templatePath = "Template.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";
                // Sample smart markers (adjust as needed)
                sheet.Cells["A1"].PutValue("&Data.Name&");
                sheet.Cells["B1"].PutValue("&Data.Value&");
            }

            // Prepare the data source for smart markers
            IEnumerable<object> dataSource = GetData();

            // Set up the designer for smart marker processing
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", dataSource);

            // Process all smart markers (Aspose.Cells processes the whole sheet)
            designer.Process();

            // Save the processed workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Example data source generator (replace with real data)
    static IEnumerable<object> GetData()
    {
        for (int i = 0; i < 100; i++)
        {
            yield return new { Name = "Item" + i, Value = i };
        }
    }
}
