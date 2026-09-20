// Title: Set DbConnection.CommandTimeout to zero for unlimited execution of long‑running queries in C#
// AI Prompts: Write C# code that opens an ADO.NET DbConnection, assigns 0 to its CommandTimeout property, runs a heavy SELECT statement, loads the result into a DataTable, and then uses Aspose.Cells to write the data to an Excel workbook. | Show how to adapt an existing Aspose.Cells workbook‑creation routine so that it receives a DbCommand whose timeout is set to zero and exports the query result without hitting a timeout error.
// Common Searches: how to set CommandTimeout to zero for long running query in C# | Aspose.Cells export data without timeout limit | C# infinite query timeout when exporting SQL data to Excel | execute large SELECT without timing out using ADO.NET | example of unlimited command timeout for DataTable fill in .NET
// Tags: C# ADO.NET unlimited command timeout | Aspose.Cells export with no query timeout | set command timeout zero .NET | long running SQL query Excel export | configure infinite timeout for DbCommand

using System;
using System.IO;
using Aspose.Cells;

// The sample program checks whether a template Excel file exists; if not, it creates a new workbook, adds a worksheet named "Data" with a simple header and one data row, auto‑fits the columns, and saves the workbook as "Result.xlsx", handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            string templatePath = "Template.xlsx";
            Workbook workbook;

            // Load existing workbook if the template file exists; otherwise create a new one
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Add sample header and data
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("John Doe");
            }

            // Auto‑fit columns for better appearance
            workbook.Worksheets[0].AutoFitColumns();

            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
