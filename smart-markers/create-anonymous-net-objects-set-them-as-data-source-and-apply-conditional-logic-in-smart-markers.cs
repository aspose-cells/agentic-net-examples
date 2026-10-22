// Title: Create anonymous .NET objects, bind them as a data source, and use IF/ELSE smart markers in Aspose.Cells (C#)
// AI Prompts: Write C# code that builds a List of anonymous objects with Name and Score properties, assigns it to WorkbookDesigner via SetDataSource, and defines a smart marker using &IF{Score>80}Pass&ELSEFail&ENDIF to generate Pass/Fail values. | Refactor the given DataTable example to use a collection of anonymous objects while preserving the conditional smart marker that checks if Score > 80.
// Common Searches: how to use a List of anonymous objects as a data source for Aspose.Cells smart markers in C# | aspnet cells smart marker IF condition with anonymous objects | binding anonymous object collection to WorkbookDesigner for Excel generation | C# example of &IF{Score>80} smart marker with Aspose.Cells | populate Excel template using smart markers and anonymous data source
// Tags: anonymous objects data source Aspose.Cells | WorkbookDesigner SetDataSource list | IF ELSE smart marker C# | conditional smart markers Aspose.Cells | populate Excel with smart markers C#

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

// The example demonstrates how to create a collection of anonymous .NET objects containing Name and Score fields, bind this collection to Aspose.Cells WorkbookDesigner using SetDataSource, and apply an IF/ELSE smart marker (&IF{Score>80}Pass&ELSEFail&ENDIF) to output Pass or Fail based on each Score. The processed workbook is then saved as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (empty template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header cells
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["C1"].PutValue("Result");

            // Insert smart markers with conditional logic
            // & = placeholder for field value
            // &IF{condition}...&ELSE...&ENDIF for conditional output
            sheet.Cells["A2"].PutValue("&=Name");
            sheet.Cells["B2"].PutValue("&=Score");
            sheet.Cells["C2"].PutValue("&IF{Score>80}Pass&ELSEFail&ENDIF");

            // Prepare data source as a DataTable (compatible with WorkbookDesigner)
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Score", typeof(int));

            dt.Rows.Add("Alice", 92);
            dt.Rows.Add("Bob", 76);
            dt.Rows.Add("Charlie", 85);

            // Bind the data source to the smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // Generates the final output based on smart markers

            // Save the populated workbook
            workbook.Save("SmartMarkersResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
