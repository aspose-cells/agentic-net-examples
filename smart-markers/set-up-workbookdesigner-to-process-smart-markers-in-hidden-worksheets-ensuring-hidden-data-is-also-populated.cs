// Title: Populate hidden worksheets with smart markers using WorkbookDesigner in Aspose.Cells for .NET
// Common Searches: how to fill smart markers in hidden Excel sheets with Aspose.Cells C# | WorkbookDesigner process hidden worksheets without losing visibility | Aspose.Cells update references after processing smart markers in hidden tabs | C# make hidden worksheets visible temporarily for smart marker processing | populate hidden worksheet data from a DataSet using Aspose.Cells
// Tags: WorkbookDesigner hidden sheet processing | smart markers visibility handling | update references after smart marker processing | C# temporary worksheet visibility Aspose.Cells | populate hidden worksheets from DataSet

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerHiddenSheetsDemo
{
    // The example loads a template workbook, creates a WorkbookDesigner with UpdateReference enabled, sets a DataSet as the data source, temporarily makes all hidden worksheets visible so their smart markers are processed, calls Process(true) to populate data, restores the original sheet visibility, and saves the resulting workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the template workbook that contains smart markers in hidden worksheets
            string templatePath = "TemplateWithHiddenSheets.xlsx";
            // Path where the processed workbook will be saved
            string outputPath = "ProcessedWithHiddenSheets.xlsx";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file not found at '{Path.GetFullPath(templatePath)}'.");
                return;
            }

            try
            {
                // Load the template workbook
                Workbook templateWorkbook = new Workbook(templatePath);

                // Create a WorkbookDesigner and assign the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = templateWorkbook,
                    // Ensure that references in other worksheets are updated after processing
                    UpdateReference = true
                };

                // -----------------------------------------------------------------
                // Prepare sample data source (replace with real data as needed)
                // -----------------------------------------------------------------
                DataSet dataSet = new DataSet();

                // Example table "Sales" that matches smart markers like &Sales.Year, &Sales.Amount
                DataTable salesTable = new DataTable("Sales");
                salesTable.Columns.Add("Year", typeof(string));
                salesTable.Columns.Add("Amount", typeof(double));

                // Populate the table with sample rows
                salesTable.Rows.Add("2020", 1250.75);
                salesTable.Rows.Add("2021", 1580.00);
                salesTable.Rows.Add("2022", 1725.30);

                dataSet.Tables.Add(salesTable);

                // Set the DataSet as the data source for the designer
                designer.SetDataSource(dataSet);

                // -----------------------------------------------------------------
                // Make hidden worksheets temporarily visible so that smart markers
                // inside them are processed. After processing we will restore the
                // original visibility state.
                // -----------------------------------------------------------------
                bool[] originalVisibility = new bool[designer.Workbook.Worksheets.Count];
                for (int i = 0; i < designer.Workbook.Worksheets.Count; i++)
                {
                    Worksheet ws = designer.Workbook.Worksheets[i];
                    originalVisibility[i] = ws.IsVisible;
                    if (!ws.IsVisible)
                    {
                        ws.IsVisible = true;
                    }
                }

                // Process all smart markers (including those in the now-visible hidden sheets)
                // The boolean parameter indicates that references to other worksheets should be updated.
                designer.Process(true);

                // Restore the original visibility of each worksheet
                for (int i = 0; i < designer.Workbook.Worksheets.Count; i++)
                {
                    designer.Workbook.Worksheets[i].IsVisible = originalVisibility[i];
                }

                // Save the processed workbook
                designer.Workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook processed successfully. Output saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
