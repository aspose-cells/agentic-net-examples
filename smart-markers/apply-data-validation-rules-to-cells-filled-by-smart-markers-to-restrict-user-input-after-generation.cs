using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook (template) ----------
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define headers
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");

                // Insert smart markers that will be replaced by data source values
                cells["A2"].PutValue("&=Persons.Name");
                cells["B2"].PutValue("&=Persons.Age");

                // Name the range that contains smart markers (required for processing)
                Aspose.Cells.Range smRange = cells.CreateRange("A2:B2");
                smRange.Name = "_CellsSmartMarkers";

                // ---------- Prepare data source ----------
                // Using an ArrayList of simple objects
                ArrayList persons = new ArrayList
                {
                    new Person { Name = "John", Age = 28 },
                    new Person { Name = "Jane", Age = 34 },
                    new Person { Name = "Bob", Age = 45 }
                };

                // ---------- Process smart markers ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = wb
                };
                designer.SetDataSource("Persons", persons);
                // Process only the defined range (true = preserve unrecognized markers)
                designer.Process(smRange, true);

                // ---------- Apply data validation to the populated Age column ----------
                // Determine the last row that contains data after processing
                int lastDataRow = sheet.Cells.MaxDataRow; // includes header row
                // Validation range: B2 to B{lastDataRow}
                CellArea ageArea = CellArea.CreateCellArea(1, 1, lastDataRow, 1); // rows are zero‑based

                ValidationCollection validations = sheet.Validations;
                int validationIndex = validations.Add(ageArea);
                Validation ageValidation = validations[validationIndex];

                // Restrict to whole numbers between 0 and 120
                ageValidation.Type = ValidationType.WholeNumber;
                ageValidation.Operator = OperatorType.Between;
                ageValidation.Formula1 = "0";
                ageValidation.Formula2 = "120";

                // Optional UI messages
                ageValidation.InputTitle = "Age Input";
                ageValidation.InputMessage = "Enter an age between 0 and 120.";
                ageValidation.ErrorTitle = "Invalid Age";
                ageValidation.ErrorMessage = "The age must be a whole number within the allowed range.";
                ageValidation.ShowInput = true;
                ageValidation.ShowError = true;
                ageValidation.InCellDropDown = false;
                ageValidation.IgnoreBlank = true;

                // ---------- Save the result ----------
                string outputPath = "SmartMarkerWithValidation.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Simple POCO used as data source for smart markers
        public class Person
        {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
        }
    }
}