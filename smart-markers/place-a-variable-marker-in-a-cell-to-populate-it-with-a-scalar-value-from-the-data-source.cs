// Title: Insert a scalar smart marker into an Excel cell and populate it from a DataSet using Aspose.Cells for .NET (C#)
// AI Prompts: Add a smart marker '&=Value' to cell A1, bind a DataSet that has a single column named 'Value' with an integer, run WorkbookDesigner.Process(), and save the workbook as Result.xlsx. | Generate C# code that creates a new Workbook, places a scalar smart marker, sets a DataSet as the data source for WorkbookDesigner, processes the markers, and writes the file to disk.
// Common Searches: Aspose.Cells C# smart marker scalar value from DataSet example | How to bind a DataSet with one column to a WorkbookDesigner smart marker | Insert variable marker '&=Value' into Excel using Aspose.Cells .NET | Populate Excel cell with integer from DataSet using smart markers in C#
// Tags: smart-marker scalar insertion Aspose.Cells | WorkbookDesigner bind DataSet C# | populate Excel cell from DataSet Aspose | C# Aspose.Cells variable marker example | process smart markers .NET workbook

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, puts the smart marker '&=Value' in cell A1, builds a DataSet with a single integer column 'Value' containing 12345, binds the DataSet to a WorkbookDesigner, processes the marker, and saves the result as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Insert a smart marker that will be replaced by the scalar value
                workbook.Worksheets[0].Cells["A1"].PutValue("&=Value");

                // Prepare a DataSet with a single column "Value" containing the scalar
                DataTable table = new DataTable();
                table.Columns.Add("Value", typeof(int));
                table.Rows.Add(12345); // scalar value to populate the marker

                DataSet dataSource = new DataSet();
                dataSource.Tables.Add(table);

                // Bind the data source to the workbook designer and process markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSource);
                designer.Process();

                // Save the resulting workbook
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
