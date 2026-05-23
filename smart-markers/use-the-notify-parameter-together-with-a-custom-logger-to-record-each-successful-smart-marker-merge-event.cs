using System;
using System.Data;
using Aspose.Cells;

// Simple logger that writes messages to the console
class SmartMarkerLogger
{
    public void Log(string message)
    {
        Console.WriteLine(message);
    }
}

// Callback implementation that logs each smart‑marker merge event
class SmartMarkerCallback : ISmartMarkerCallBack
{
    private readonly SmartMarkerLogger _logger;

    public SmartMarkerCallback(SmartMarkerLogger logger)
    {
        _logger = logger;
    }

    // This method is invoked by WorkbookDesigner for every smart marker it processes
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        _logger.Log($"Smart marker merged – Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
    }
}

class Program
{
    static void Main()
    {
        // ---------- Create a workbook with smart markers ----------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert sample smart markers
        sheet.Cells["A1"].PutValue("&=Employees.Name");
        sheet.Cells["B1"].PutValue("&=Employees.Age");

        // ---------- Prepare a data source ----------
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Age", typeof(int));
        employees.Rows.Add("Alice", 30);
        employees.Rows.Add("Bob", 28);

        // ---------- Set up logger and callback ----------
        SmartMarkerLogger logger = new SmartMarkerLogger();
        SmartMarkerCallback callback = new SmartMarkerCallback(logger);

        // ---------- Configure WorkbookDesigner ----------
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;          // assign workbook
        designer.CallBack = callback;          // assign custom callback
        designer.SetDataSource(employees);     // set data source

        // ---------- Process smart markers ----------
        // The callback will be triggered for each successful merge
        designer.Process();

        // ---------- Save the result ----------
        workbook.Save("SmartMarkerOutput.xlsx");
    }
}