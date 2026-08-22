// Title: Check that a cell's string length stays unchanged after disabling Excel 2003 compatibility and refreshing a pivot table with Aspose.Cells for .NET
// AI Prompts: Generate C# code that writes a 300‑character string to a worksheet, creates a pivot table on that range, sets IsExcel2003Compatible = false, refreshes the pivot cache, saves the workbook as .xlsx using OoxmlSaveOptions, reloads it, and compares the original and loaded cell lengths. | Show how to open a saved workbook with Aspose.Cells LoadOptions and verify that long text in the source cell is not truncated after the pivot table refresh. | Provide a console output snippet that prints the original cell length, the length after reload, and a boolean indicating whether they match.
// Common Searches: aspnet aspose.cells verify long text length after pivot table refresh | disable Excel2003 compatibility pivot table Aspose.Cells .NET | preserve >255 characters in pivot cache Aspose.Cells example | how to check cell string length after saving workbook with Aspose.Cells | Aspose.Cells pivot table refresh does not truncate long strings
// Tags: pivot table refresh without truncation Aspose.Cells | excel2003 compatibility false Aspose.Cells | preserve long string length .xlsx Aspose.Cells | verify cell content length after workbook reload | use OoxmlSaveOptions with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, inserts a 300‑character string into a cell, builds a pivot table on that data, disables Excel 2003 compatibility, refreshes and calculates the pivot, saves the file as .xlsx, reloads it, and compares the original and loaded cell string lengths to confirm they are identical, then removes the temporary file.
    public class VerifyCellContentLengthAfterCompatibilityDisable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and add sample data
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header
                dataSheet.Cells["A1"].PutValue("LongText");

                // Create a string longer than Excel 2003 limit (255 characters)
                string longString = new string('x', 300);
                dataSheet.Cells["A2"].PutValue(longString);

                // Record original length of the cell content
                int originalLength = dataSheet.Cells["A2"].StringValue.Length;

                // Add a second worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Add a pivot table based on the data range A1:A2
                int pivotIndex = pivotSheet.PivotTables.Add("PivotTable", "A1:A2", "A4", false);
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add the long text column as a row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

                // Disable Excel 2003 compatibility to avoid truncation
                pivotTable.IsExcel2003Compatible = false;

                // Refresh pivot cache data (correct API)
                pivotTable.RefreshData();

                // Calculate the pivot table data
                pivotTable.CalculateData();

                // Save the workbook using OoxmlSaveOptions
                string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                workbook.Save(tempFile, saveOptions);

                // Load the workbook back if the file exists
                if (File.Exists(tempFile))
                {
                    try
                    {
                        LoadOptions loadOptions = new LoadOptions();
                        Workbook loadedWorkbook = new Workbook(tempFile, loadOptions);

                        // Retrieve the length of the cell content after reload
                        Worksheet loadedDataSheet = loadedWorkbook.Worksheets["Data"];
                        int loadedLength = loadedDataSheet.Cells["A2"].StringValue.Length;

                        // Verify that the length remains unchanged
                        Console.WriteLine("Original length: " + originalLength);
                        Console.WriteLine("Loaded length:   " + loadedLength);
                        Console.WriteLine("Length unchanged: " + (originalLength == loadedLength));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error loading workbook: " + ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("Temporary file was not created.");
                }

                // Clean up temporary file
                if (File.Exists(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error deleting temporary file: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
