// Title: Configure Aspose.Cells WorkbookDesigner to ignore empty smart markers and avoid extra rows in C#
// Common Searches: how to stop Aspose.Cells WorkbookDesigner from adding rows for empty smart markers in C# | ignore empty smart marker collections when processing a workbook with Aspose.Cells | skip creating rows for smart markers with no data source in Aspose.Cells C# | prevent extra rows in Excel output when smart marker dataset is empty using Aspose.Cells | Aspose.Cells WorkbookDesigner handling empty DataSet for smart markers
// Tags: WorkbookDesigner empty smart marker handling | skip empty rows Aspose.Cells smart markers | prevent extra rows Aspose.Cells workbook | C# Aspose.Cells smart marker empty dataset | smart markers ignore empty collections

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook template containing smart markers, creates an empty DataSet with a DataTable that has no rows, assigns it to a WorkbookDesigner, processes the smart markers, and saves the result. It demonstrates how to configure WorkbookDesigner to ignore empty smart markers and prevent unnecessary row creation when the source collection is empty.
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = "TemplateWithSmartMarkers.xlsx";
            string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            try
            {
                // Load the workbook from the template file
                Workbook workbook = new Workbook(templatePath);

                // Initialize the WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Prepare an empty data source (DataSet with an empty DataTable)
                DataSet dataSet = new DataSet();
                DataTable itemsTable = new DataTable("Items");
                itemsTable.Columns.Add("Value", typeof(string));
                // No rows are added, keeping the table empty
                dataSet.Tables.Add(itemsTable);

                // Set the data source for the designer
                designer.SetDataSource(dataSet);

                // Process the smart markers in the workbook
                designer.Process();

                // Save the resulting workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Result workbook saved to: {resultPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
