// Title: Configure a DBConnection to have an unlimited CommandTimeout in C#
// Common Searches: how to set CommandTimeout to 0 for SqlConnection in C# | remove query timeout for OleDbConnection C# long running query | C# execute stored procedure without timeout limit | set unlimited command timeout for ADO.NET database connection | prevent timeout errors when running large data import in C#
// Tags: set CommandTimeout zero C# | unlimited query execution ADO.NET | disable database command timeout | configure DBConnection timeout property | run long-running SQL command without timeout

using System;
using System.IO;
using Aspose.Cells;

// Shows how to assign zero to the CommandTimeout property of a DBConnection (such as SqlConnection or OleDbConnection) in C#. A zero value disables the timeout, allowing queries that take an indefinite amount of time to complete without raising timeout exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the Excel file (template or new)
            string excelPath = "Template.xlsx";

            Workbook workbook;

            // Load existing workbook if file exists, otherwise create a new one
            if (File.Exists(excelPath))
            {
                workbook = new Workbook(excelPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Write some data to cell A1
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
