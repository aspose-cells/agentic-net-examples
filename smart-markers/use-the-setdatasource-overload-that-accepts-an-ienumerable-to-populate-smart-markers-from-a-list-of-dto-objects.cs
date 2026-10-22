// Title: Bind a List of PersonDto objects to Aspose.Cells smart markers with WorkbookDesigner.SetDataSource(IEnumerable) in C#
// AI Prompts: Generate C# code that loads an Excel template, creates a List<PersonDto>, sets it as a smart‑marker data source using WorkbookDesigner.SetDataSource("persons", persons), processes the markers, recalculates formulas, and saves the workbook. | Show how to replace the DTO class and data source identifier in the smart‑marker example while still using the IEnumerable overload of SetDataSource. | Explain step‑by‑step how Aspose.Cells WorkbookDesigner can populate smart markers from any IEnumerable collection and then trigger formula recalculation.
// Common Searches: aspnet aspocells setdatasource ienumerable smart markers example | how to populate smart markers from a List<T> using Aspose.Cells in C# | WorkbookDesigner SetDataSource with custom DTO collection | recalculate formulas after processing smart markers Aspose.Cells | using Aspose.Cells smart markers with a list of objects in .NET
// Tags: WorkbookDesigner SetDataSource with IEnumerable | Aspose.Cells smart markers using DTO list | populate smart markers from List<T> | recalculate workbook formulas Aspose.Cells | process Excel template smart markers C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // DTO class representing the data for smart markers
    // The example loads an Excel template, creates a List<PersonDto>, binds it to smart markers via WorkbookDesigner.SetDataSource("persons", persons), processes the markers, recalculates all formulas, and saves the populated workbook.
    public class PersonDto
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public double Salary { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                const string templatePath = "SmartMarkerTemplate.xlsx";
                const string resultPath = "SmartMarkerResult.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");

                // Load a workbook that contains smart markers (template)
                Workbook workbook = new Workbook(templatePath);

                // Create a list of DTO objects to serve as the data source
                List<PersonDto> persons = new List<PersonDto>
                {
                    new PersonDto { Name = "John", Age = 30, Salary = 50000 },
                    new PersonDto { Name = "Jane", Age = 28, Salary = 55000 },
                    new PersonDto { Name = "Bob", Age = 35, Salary = 60000 }
                };

                // Initialize the WorkbookDesigner (smart marker processor) and set the data source
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                // Provide a name for the data source collection (e.g., "persons")
                designer.SetDataSource("persons", persons);
                designer.Process(); // Populate smart markers

                // Recalculate formulas after processing smart markers
                workbook.CalculateFormula();

                // Save the workbook with the populated data
                workbook.Save(resultPath);
                Console.WriteLine($"Smart marker processing completed. Result saved to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
