// Title: Transposing a DataTable and filling a smart marker with the Range parameter using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells example to transpose a DataTable for smart marker import | How to use the Range parameter to swap rows and columns in a smart marker | WorkbookDesigner processing of transposed data for smart markers in .NET | Fill a smart marker layout with transposed DataTable using Aspose.Cells | Smart marker row‑column swap tutorial with Aspose.Cells for C#
// Tags: WorkbookDesigner transposed data source | smart marker layout using range | C# DataTable transposition with Aspose.Cells | row column swapping in smart markers | Aspose.Cells fill smart marker from transposed table

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds a smart marker with a Range parameter, builds a DataTable, transposes it so rows become columns, sets the transposed table as the data source for WorkbookDesigner, processes the smart marker, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add a template worksheet.
            Workbook workbook = new Workbook();
            Worksheet templateSheet = workbook.Worksheets[0];
            templateSheet.Name = "Template";

            // 2. Insert a smart marker with a *Range* parameter.
            //    The syntax &="Data" tells the designer to import the data source named "Data".
            templateSheet.Cells["A1"].PutValue("&=\"Data\"");

            // 3. Build a sample DataTable that will be used as the data source.
            DataTable dataTable = new DataTable("Data");
            dataTable.Columns.Add("Product", typeof(string));
            dataTable.Columns.Add("Jan", typeof(double));
            dataTable.Columns.Add("Feb", typeof(double));
            dataTable.Columns.Add("Mar", typeof(double));

            dataTable.Rows.Add("Apple", 120, 130, 140);
            dataTable.Rows.Add("Banana", 80, 90, 100);
            dataTable.Rows.Add("Cherry", 50, 60, 70);

            // 4. Transpose the DataTable to achieve row‑column swapping.
            DataTable transposedTable = TransposeDataTable(dataTable);

            // 5. Set the data source for the smart marker and process it.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", transposedTable);
            designer.Process();

            // 6. Save the resulting workbook.
            string outputPath = "SmartMarkerSwapRowsColumns.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Helper method to transpose a DataTable (rows become columns and vice‑versa).
    private static DataTable TransposeDataTable(DataTable source)
    {
        DataTable result = new DataTable();

        // First column will hold the original column names.
        result.Columns.Add(source.Columns[0].ColumnName, source.Columns[0].DataType);

        // Add a column for each row in the source table.
        foreach (DataRow row in source.Rows)
        {
            string newColumnName = row[0].ToString();
            result.Columns.Add(newColumnName, typeof(string));
        }

        // Populate rows with the original column names (except the first column).
        for (int i = 1; i < source.Columns.Count; i++)
        {
            DataRow newRow = result.NewRow();
            newRow[0] = source.Columns[i].ColumnName;
            for (int j = 0; j < source.Rows.Count; j++)
            {
                newRow[j + 1] = source.Rows[j][i];
            }
            result.Rows.Add(newRow);
        }

        return result;
    }
}
