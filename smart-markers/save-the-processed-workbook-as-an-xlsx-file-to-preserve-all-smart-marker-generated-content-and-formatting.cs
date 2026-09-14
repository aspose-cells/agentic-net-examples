// Title: Save a workbook with processed smart markers to XLSX while retaining all generated data and formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a workbook, places smart markers, assigns a List<Person> as the data source, executes WorkbookDesigner.Process, and writes the result to an XLSX file while keeping all cell formatting. | Show the steps to export the output of Aspose.Cells smart marker processing to an .xlsx document without losing any generated values or visual formatting in a .NET project.
// Common Searches: how to retain smart marker generated values when saving to .xlsx using Aspose.Cells in C# | C# Aspose.Cells save workbook after smart marker processing with original formatting | example code for exporting processed smart markers to an Excel file without formatting loss | Aspose.Cells WorkbookDesigner Process save as XLSX preserving styles
// Tags: Aspose.Cells WorkbookDesigner Process to XLSX | smart markers export with formatting | C# save processed workbook as XLSX | preserve generated smart marker content | binding collection to smart markers Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

// The program creates a workbook, defines smart markers, binds a list of Person objects as the data source, processes the markers with WorkbookDesigner, and saves the resulting workbook as an XLSX file, preserving all generated content and cell formatting.
class Program
{
    static void Main()
    {
        // Create a new workbook (template)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define smart markers in the worksheet
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("&=$Name");
        sheet.Cells["B2"].PutValue("&=$Age");

        // Prepare a data source for the smart markers
        ArrayList persons = new ArrayList();
        persons.Add(new Person { Name = "John Doe", Age = 30 });
        persons.Add(new Person { Name = "Jane Smith", Age = 25 });

        // Set the data source and process the smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Data", persons);
        designer.Process();

        // Save the processed workbook as XLSX, preserving all generated content and formatting
        workbook.Save("SmartMarkersResult.xlsx", SaveFormat.Xlsx);
    }

    // Simple POCO class used as a data source
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
