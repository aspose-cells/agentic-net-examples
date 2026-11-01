// Title: Use Aspose.Cells WorkbookDesigner in C# to process smart markers and save the workbook as XLSX with formatting preserved
// AI Prompts: Write C# code that loads an Excel file containing smart markers, binds a DataSet to WorkbookDesigner, processes the markers, and saves the result as an XLSX file while keeping all formatting intact. | Show how to populate a DataSet with employee data, apply it to smart markers using Aspose.Cells WorkbookDesigner, and export the processed workbook to a new .xlsx file.
// Common Searches: Aspose.Cells C# bind DataSet to smart markers and export to xlsx | preserve formatting when saving processed smart markers with WorkbookDesigner | C# example for processing smart markers and saving workbook as XLSX using Aspose.Cells | save smart marker populated Excel file as .xlsx with Aspose.Cells .NET | WorkbookDesigner process smart markers and keep generated content in output.xlsx
// Tags: WorkbookDesigner process smart markers | save workbook as XLSX Aspose.Cells | bind DataSet to smart markers C# | preserve formatting Aspose.Cells XLSX export | smart markers data source Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// The sample loads an input.xlsx workbook that contains smart markers, creates a DataSet with an Employees table, assigns the DataSet to WorkbookDesigner, processes the smart markers, and saves the resulting workbook as output.xlsx in XLSX format, preserving all generated content and formatting.
class Program
{
    static void Main()
    {
        // Load the workbook that contains smart markers.
        Workbook workbook = new Workbook("input.xlsx");

        // Prepare the data source for the smart markers.
        // In a real scenario, populate the DataSet with the actual data.
        DataSet dataSet = new DataSet();

        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Age", typeof(int));

        employees.Rows.Add("John Doe", 30);
        employees.Rows.Add("Jane Smith", 28);
        dataSet.Tables.Add(employees);

        // Process the smart markers using WorkbookDesigner.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dataSet);
        designer.Process();

        // Save the processed workbook as XLSX, preserving all generated content and formatting.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
