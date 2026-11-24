// Title: Load an XLSX workbook from a stream with cache parsing, modify the first pivot table (add a data field and move a column to rows), and save back to a stream using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook from memory stream with cache parsing option | how to add a data field to a pivot table using Aspose.Cells in .NET | move a column field to rows in an Aspose.Cells pivot table programmatically | save modified Excel workbook to output stream with Aspose.Cells OoxmlSaveOptions | process pivot table from input stream and write result to output stream using Aspose.Cells
// Tags: cache parsing load options Aspose.Cells | add Quantity data field to pivot Aspose.Cells C# | transfer first column field to row area pivot Aspose.Cells | recalculate pivot data Aspose.Cells | save workbook to output stream OoxmlSaveOptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example demonstrates loading an XLSX workbook from an input stream with cache parsing enabled, locating the first pivot table, adding a "Quantity" data field, moving the first column field to the row area when no row fields exist, recalculating the pivot, and saving the updated workbook to an output stream using OoxmlSaveOptions, while handling errors and resetting stream positions.
public class PivotModifier
{
    public void ProcessPivot(Stream inputStream, Stream outputStream)
    {
        try
        {
            // Load the workbook from the input stream
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputStream, loadOptions);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
                throw new InvalidOperationException("The workbook contains no worksheets.");

            // Find the first pivot table in the workbook
            PivotTable pivot = null;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.PivotTables.Count > 0)
                {
                    pivot = sheet.PivotTables[0];
                    break;
                }
            }

            if (pivot == null)
                throw new InvalidOperationException("No pivot table found in the workbook.");

            // Add a new data field named "Quantity"
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // If there are no row fields but column fields exist, move the first column field to rows
            if (pivot.RowFields.Count == 0 && pivot.ColumnFields.Count > 0)
            {
                var columnField = pivot.ColumnFields[0];
                string fieldName = columnField.Name;

                // Add the field to the row area
                pivot.AddFieldToArea(PivotFieldType.Row, fieldName);

                // Remove the field from the column area
                pivot.ColumnFields.RemoveAt(0);
            }

            // Recalculate pivot data after modifications
            pivot.CalculateData();

            // Save the modified workbook to the output stream
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
            workbook.Save(outputStream, saveOptions);
        }
        catch (Exception ex)
        {
            // Log the error and rethrow for the caller to handle if needed
            Console.Error.WriteLine($"Error processing pivot: {ex.Message}");
            throw;
        }
        finally
        {
            // Reset stream positions for potential further use
            if (inputStream.CanSeek) inputStream.Position = 0;
            if (outputStream.CanSeek) outputStream.Position = 0;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists before proceeding
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                var modifier = new PivotModifier();
                modifier.ProcessPivot(inputStream, outputStream);
            }

            Console.WriteLine($"Pivot table modified and saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
