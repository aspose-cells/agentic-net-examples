// Title: Use Aspose.Cells for .NET to apply conditional smart markers that display “Pass” or “Fail” based on a numeric Score field
// Common Searches: how to insert an IF smart marker in Aspose.Cells C# workbook | Aspose.Cells conditional smart marker syntax for numeric comparison | C# example of using WorkbookDesigner with DataSet to generate Pass/Fail results | display Pass or Fail in Excel using Aspose.Cells smart markers and a Score column | process conditional smart markers in Aspose.Cells after setting a DataSource
// Tags: conditional smart marker IF syntax Aspose.Cells | WorkbookDesigner SetDataSource DataSet C# | generate Pass Fail text with smart markers | Aspose.Cells smart marker processing workflow | Excel export conditional text using Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace ConditionalSmartMarkerExample
{
    // Simple data class with a numeric score property
    // The example creates a new workbook, places the conditional smart marker '&IF{Score>=60}Pass;Fail&' in cell A1, builds a DataTable from a list of Student objects, wraps it in a DataSet, assigns the DataSet to a WorkbookDesigner, processes the smart markers, and saves the resulting XLSX file.
    public class Student
    {
        public int Score { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a conditional smart marker that shows "Pass" if Score >= 60, otherwise "Fail"
                // Syntax: &IF{condition}TrueText;FalseText&
                sheet.Cells["A1"].PutValue("&IF{Score>=60}Pass;Fail&");

                // Prepare sample data
                List<Student> data = new List<Student>
                {
                    new Student { Score = 85 },
                    new Student { Score = 55 },
                    new Student { Score = 70 }
                };

                // Convert the list to a DataTable and wrap it in a DataSet (required by SetDataSource overload)
                DataTable dt = new DataTable("Students");
                dt.Columns.Add("Score", typeof(int));
                foreach (Student s in data)
                {
                    dt.Rows.Add(s.Score);
                }

                DataSet ds = new DataSet();
                ds.Tables.Add(dt);

                // Set the data source and process the smart markers (lifecycle: load/process)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(ds);
                designer.Process();

                // Save the workbook (lifecycle: save)
                string outputPath = "ConditionalSmartMarkerResult.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
