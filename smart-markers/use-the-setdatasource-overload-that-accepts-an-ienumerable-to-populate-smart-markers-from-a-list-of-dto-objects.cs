// Title: How to bind a List of PersonDto objects to Aspose.Cells smart markers using WorkbookDesigner.SetDataSource in C#
// Common Searches: C# example for binding a List<PersonDto> to Excel smart markers with Aspose.Cells | using WorkbookDesigner SetDataSource to populate smart markers from a collection | Aspose.Cells smart marker data source from custom DTO list | populate Excel template smart markers with a List of objects in .NET | convert List<T> to DataSet for Aspose.Cells smart markers
// Tags: WorkbookDesigner SetDataSource DataSet overload | binding DTO list to Aspose.Cells smart markers | C# convert List<PersonDto> to DataTable | Excel template smart markers with Aspose.Cells | Aspose.Cells data source for smart markers

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // DTO class representing the data for smart markers
    // The sample loads or creates an Excel workbook with smart marker placeholders, defines a PersonDto class, builds a List<PersonDto>, converts the list into a DataTable and DataSet, sets the DataSet as the data source for WorkbookDesigner, processes the smart markers, and saves the populated workbook.
    public class PersonDto
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Load existing template or create a minimal workbook with smart markers
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("&=Name");
                    sheet.Cells["B1"].PutValue("&=Age");
                }

                // Prepare data source as a DataSet (required by WorkbookDesigner.SetDataSource)
                List<PersonDto> people = new List<PersonDto>
                {
                    new PersonDto { Name = "John", Age = 30 },
                    new PersonDto { Name = "Jane", Age = 25 },
                    new PersonDto { Name = "Bob",  Age = 40 }
                };

                DataTable table = new DataTable("People");
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Age", typeof(int));

                foreach (PersonDto person in people)
                {
                    table.Rows.Add(person.Name, person.Age);
                }

                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(table);

                // Bind data and process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process();

                // Save the processed workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
