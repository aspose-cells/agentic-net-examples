// Title: C# post‑processing callback to modify Excel cells after Aspose.Cells smart marker processing before saving
// Common Searches: how to execute custom code after WorkbookDesigner.Process in Aspose.Cells C# | modify Excel column values after smart marker evaluation using Aspose.Cells | Aspose.Cells C# example of post‑processing callback for smart markers | prepend text to cells in column A after smart marker processing with Aspose.Cells
// Tags: smart marker post‑processing callback C# | modify cell values after WorkbookDesigner.Process | prepend text to Excel column using Aspose.Cells | custom post‑processing of Aspose.Cells smart markers | Aspose.Cells event handling for smart markers

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook containing smart markers, sets a DataSet as the data source, processes the markers with WorkbookDesigner, then runs a custom OnAfterSmartMarkerProcessing callback that prepends "Processed-" to every value in column A of the first worksheet before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "InputWithSmartMarkers.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load an existing workbook that contains smart markers
            Workbook workbook = new Workbook(inputPath);

            // Create a WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a DataSet as the data source for the smart markers
            DataSet dataSet = new DataSet();
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));

            employeeTable.Rows.Add("John", 30);
            employeeTable.Rows.Add("Jane", 25);

            dataSet.Tables.Add(employeeTable);
            designer.SetDataSource(dataSet);

            // Process the smart markers
            designer.Process();

            // Execute post‑processing logic
            OnAfterSmartMarkerProcessing(designer, EventArgs.Empty);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Callback method executed after smart markers have been processed
    private static void OnAfterSmartMarkerProcessing(object sender, EventArgs e)
    {
        try
        {
            if (sender is WorkbookDesigner designer)
            {
                Workbook wb = designer.Workbook;

                // Example modification: prepend "Processed-" to all values in column A of the first worksheet
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                int maxRow = cells.MaxDataRow;
                for (int row = 0; row <= maxRow; row++)
                {
                    Cell cell = cells[row, 0]; // Column A (index 0)
                    if (cell.Value != null)
                    {
                        cell.PutValue("Processed-" + cell.Value.ToString());
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Post‑processing error: {ex.Message}");
        }
    }
}
