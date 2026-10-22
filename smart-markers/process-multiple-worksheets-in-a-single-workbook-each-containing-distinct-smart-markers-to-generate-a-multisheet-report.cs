// Title: Generate a multi‑sheet Excel report in C# using Aspose.Cells smart markers with separate data sources per worksheet
// AI Prompts: Write C# code that loads an Excel template, creates a distinct DataSet for each worksheet, uses WorkbookDesigner to bind the data, processes the smart markers, and saves the final workbook. | Show how to fill Employees, Sales, and Summary sheets with different DataTables via Aspose.Cells smart markers in a single .NET application. | Explain step‑by‑step how to process smart markers sheet‑by‑sheet and produce a combined multi‑sheet Excel file.
// Common Searches: how to use Aspose.Cells WorkbookDesigner with multiple worksheets in C# | Aspose.Cells smart markers separate dataset for each sheet example | C# generate multi‑sheet Excel report from template using smart markers | populate different Excel sheets with distinct DataTables using Aspose.Cells | process smart markers per worksheet Aspose.Cells .NET
// Tags: Aspose.Cells WorkbookDesigner per-sheet data binding | smart markers multi-sheet generation C# | generate multi-sheet Excel report Aspose.Cells | assign distinct DataSet to Excel worksheet | process smart markers sheet by sheet

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook that contains smart markers, builds three separate DataSet objects for Employees, Sales, and Summary data, and uses a dedicated WorkbookDesigner instance for each worksheet to bind its data source and process the markers. After processing all sheets, the workbook is saved as a combined multi‑sheet report (MultiSheetReport.xlsx) with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "MultiSheetReport.xlsx";

            // Verify that the template file exists.
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the template workbook containing smart markers.
            Workbook workbook = new Workbook(templatePath);

            // -------------------------------------------------
            // Prepare distinct data sources for each worksheet.
            // -------------------------------------------------

            // ----- Employees worksheet data -----
            DataSet employeeDataSet = new DataSet();
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Columns.Add("Department", typeof(string));
            employeeTable.Rows.Add("John Doe", 30, "HR");
            employeeTable.Rows.Add("Jane Smith", 27, "Finance");
            employeeTable.Rows.Add("Bob Johnson", 45, "IT");
            employeeDataSet.Tables.Add(employeeTable);

            // ----- Sales worksheet data -----
            DataSet salesDataSet = new DataSet();
            DataTable salesTable = new DataTable("Sales");
            salesTable.Columns.Add("Product", typeof(string));
            salesTable.Columns.Add("Quantity", typeof(int));
            salesTable.Columns.Add("Revenue", typeof(int));
            salesTable.Rows.Add("Laptop", 120, 150000);
            salesTable.Rows.Add("Smartphone", 340, 255000);
            salesTable.Rows.Add("Tablet", 210, 105000);
            salesDataSet.Tables.Add(salesTable);

            // ----- Summary worksheet data -----
            DataSet summaryDataSet = new DataSet();
            DataTable summaryTable = new DataTable("Summary");
            summaryTable.Columns.Add("ReportDate", typeof(DateTime));
            summaryTable.Columns.Add("TotalEmployees", typeof(int));
            summaryTable.Columns.Add("TotalRevenue", typeof(int));
            summaryTable.Rows.Add(DateTime.Today, 3, 510000);
            summaryDataSet.Tables.Add(summaryTable);

            // -------------------------------------------------
            // Process each worksheet individually with its own data source.
            // -------------------------------------------------

            // Process Employees sheet (index 0).
            WorkbookDesigner designer0 = new WorkbookDesigner(workbook);
            designer0.SetDataSource(employeeDataSet);
            designer0.Process();

            // Process Sales sheet (index 1).
            WorkbookDesigner designer1 = new WorkbookDesigner(workbook);
            designer1.SetDataSource(salesDataSet);
            designer1.Process();

            // Process Summary sheet (index 2).
            WorkbookDesigner designer2 = new WorkbookDesigner(workbook);
            designer2.SetDataSource(summaryDataSet);
            designer2.Process();

            // -------------------------------------------------
            // Save the generated multi‑sheet report.
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the error.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
