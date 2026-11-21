// Title: Calculate Age with a Smart Marker Using the Formula Parameter and DATEDIF in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells smart marker formula to compute age from birthdate | How to use DATEDIF in Aspose.Cells smart markers for age calculation | Aspose.Cells WorkbookDesigner calculate age column with Formula parameter | Smart marker example for dynamic age field using Excel functions in .NET | Generate Excel report with age column using Aspose.Cells smart markers
// Tags: smart marker formula age | Aspose.Cells DATEDIF function | WorkbookDesigner data source DataTable | excel age column generation | csharp Aspose.Cells smart marker example

using System;
using System.Data;
using Aspose.Cells;

// This C# example demonstrates how to add a smart marker that calculates a person's age from a birthdate using the Formula parameter with Excel's DATEDIF function. It creates a workbook, defines header cells (Name, BirthDate, Age), inserts smart markers (including {Age:Formula=DATEDIF({BirthDate},TODAY(),"Y")} ), populates a DataTable as the data source, processes the markers with WorkbookDesigner, and saves the output as an XLSX file.
class SmartMarkerAgeExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set up header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("BirthDate");
        sheet.Cells["C1"].PutValue("Age");

        // Insert smart markers in the second row
        // {Name} and {BirthDate} will be replaced by data source values
        // {Age:Formula=DATEDIF({BirthDate},TODAY(),"Y")} calculates age using Excel's DATEDIF function
        sheet.Cells["A2"].PutValue("{Name}");
        sheet.Cells["B2"].PutValue("{BirthDate}");
        sheet.Cells["C2"].PutValue("{Age:Formula=DATEDIF({BirthDate},TODAY(),\"Y\")}");

        // Prepare a data source (DataTable) with sample data
        DataTable dt = new DataTable("People");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("BirthDate", typeof(DateTime));

        dt.Rows.Add("Alice", new DateTime(1990, 5, 12));
        dt.Rows.Add("Bob", new DateTime(1985, 11, 23));
        dt.Rows.Add("Charlie", new DateTime(2000, 2, 5));

        // Apply the smart markers using WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process();

        // Save the result
        workbook.Save("SmartMarker_Age_Output.xlsx");
    }
}
