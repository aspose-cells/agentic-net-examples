using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists before loading
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the workbook template that contains markers and the desired cell formatting
            Workbook templateWorkbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded template
            WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

            // Sample data source that will be merged into the template
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John", Age = 28 },
                new Person { Name = "Anna", Age = 32 }
            };

            // Bind the data source to the designer (markers like &Person.Name, &Person.Age should exist in the template)
            designer.SetDataSource("Person", persons);

            // Process the template – this generates the records in the worksheet
            designer.Process();

            // ------------------------------------------------------------
            // Inherit cell formatting (CopyStyle) for the generated records
            // ------------------------------------------------------------
            Worksheet sheet = designer.Workbook.Worksheets[0];

            // Assume the first data row in the template (row index 1, i.e., A2) has the style we want to copy
            Style templateStyle = sheet.Cells["A2"].GetStyle();

            // Determine the range of rows that now contain data after processing
            int firstDataRow = 1; // zero‑based index of the first row with data (A2)
            int lastDataRow = sheet.Cells.MaxDataRow; // last row with data

            // Apply the template style to each populated cell in the generated rows
            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                for (int col = 0; col <= sheet.Cells.MaxDataColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Skip cells that have no value
                    if (cell.Value == null)
                        continue;

                    // Create a new style instance and copy the template style into it
                    Style newStyle = designer.Workbook.CreateStyle();
                    newStyle.Copy(templateStyle);

                    // Assign the copied style to the current cell
                    cell.SetStyle(newStyle);
                }
            }

            // Save the resulting workbook
            designer.Workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Simple POCO class used as data source
    public class Person
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }
}