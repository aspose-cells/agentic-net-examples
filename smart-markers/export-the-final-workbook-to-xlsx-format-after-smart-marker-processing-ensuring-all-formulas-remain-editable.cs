// Title: Save a workbook after smart marker processing to XLSX while keeping formulas editable using Aspose.Cells for .NET
// Common Searches: C# how to export a smart marker processed Excel file as .xlsx with formulas intact using Aspose.Cells | Aspose.Cells WorkbookDesigner export to XLSX preserving editable formulas | process smart markers and save result workbook without converting formulas to values in .NET | save smart marker template to XLSX while retaining formula cells in C#
// Tags: WorkbookDesigner smart marker processing | export workbook to xlsx Aspose.Cells | preserve editable formulas Aspose.Cells | C# DataTable smart marker data source | save processed template as xlsx

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an Excel template containing smart markers (or creates one), assigns a DataTable as the data source via WorkbookDesigner, processes the smart markers, and then saves the resulting workbook as an XLSX file, ensuring that all formulas remain editable.
class SmartMarkerExport
{
    static void Main()
    {
        try
        {
            string templatePath = "TemplateWithSmartMarkers.xlsx";
            Workbook workbook;

            // Load existing template if present; otherwise create a simple workbook with placeholder smart markers
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Name: &=\"DataSource.Name\"");
                ws.Cells["A2"].PutValue("Age: &=\"DataSource.Age\"");
            }

            // Initialize designer and set the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            DataTable dataSource = GetDataSource();
            designer.SetDataSource(dataSource); // SetDataSource accepts DataTable
            designer.Process(); // Process smart markers

            // Save the result workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Sample data source for smart markers (DataTable with columns matching marker fields)
    static DataTable GetDataSource()
    {
        DataTable dt = new DataTable("DataSource");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("John Doe", 30);
        dt.Rows.Add("Jane Smith", 25);
        return dt;
    }
}
