using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerAgeDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a template workbook with smart markers
            // -------------------------------------------------
            Workbook templateWb = new Workbook();
            Worksheet sheet = templateWb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("BirthDate");
            cells["B1"].PutValue("Age");

            // Smart marker for birthdate value
            cells["A2"].PutValue("&=[People].BirthDate");

            // Smart marker with Formula to calculate age
            // The formula uses DATEDIF to compute years between birthdate and today
            // Note: The inner smart marker placeholder is escaped by using double quotes for the string argument
            cells["B2"].PutValue("&=[People].BirthDate?Formula=DATEDIF(&=[People].BirthDate,TODAY(),\"Y\")");

            // Save the template to a memory stream (no file I/O)
            MemoryStream templateStream = new MemoryStream();
            templateWb.Save(templateStream, SaveFormat.Xlsx);
            templateStream.Position = 0; // Reset stream position for reading

            // -------------------------------------------------
            // 2. Load the template into WorkbookDesigner
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(templateStream);

            // -------------------------------------------------
            // 3. Prepare data source (DataTable with BirthDate column)
            // -------------------------------------------------
            DataTable peopleTable = new DataTable("People");
            peopleTable.Columns.Add("BirthDate", typeof(DateTime));

            // Sample data
            peopleTable.Rows.Add(new DateTime(1990, 5, 15));
            peopleTable.Rows.Add(new DateTime(1985, 12, 30));
            peopleTable.Rows.Add(new DateTime(2000, 1, 1));

            // Set the data source for the designer
            designer.SetDataSource(peopleTable);

            // Ensure formulas are calculated after processing
            designer.CalculateFormula = true;

            // -------------------------------------------------
            // 4. Process smart markers
            // -------------------------------------------------
            designer.Process();

            // -------------------------------------------------
            // 5. Save the resulting workbook to a file
            // -------------------------------------------------
            designer.Workbook.Save("PeopleWithAge.xlsx");
        }
    }
}