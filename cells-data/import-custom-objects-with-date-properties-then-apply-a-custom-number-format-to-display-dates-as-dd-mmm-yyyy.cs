// Title: Import a list of C# objects with DateTime fields into an Aspose.Cells worksheet and format the date column as dd-MMM-yyyy
// AI Prompts: Write C# code that uses Aspose.Cells to import an array of custom objects containing a DateTime property into a worksheet starting at cell A1. | Show how to create a style with the custom number format "dd-MMM-yyyy" and apply it only to the numeric format of a specific column using a style flag. | Demonstrate saving the workbook after importing the objects and applying the date format to the BirthDate column.
// Common Searches: asp.net import list of objects into Excel with Aspose.Cells and format date column | how to apply custom date format dd-MMM-yyyy to a column after ImportCustomObjects in C# | using StyleFlag to set number format only for a column in Aspose.Cells | Aspose.Cells import custom class with DateTime property and preserve date formatting | C# Aspose.Cells example for importing objects and formatting dates
// Tags: import custom objects Aspose.Cells C# | date column format dd-MMM-yyyy Aspose.Cells | column number format styling Aspose.Cells | preserve DateTime formatting after ImportCustomObjects | export workbook to XLSX using Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The program creates a new workbook, imports an array of Person objects (including a BirthDate DateTime field) into the first worksheet, defines a style with the custom date format dd-MMM-yyyy, applies the style only to the number format of the BirthDate column, and saves the workbook as CustomObjectsWithFormattedDates.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a list of custom objects that contain a DateTime property
            List<Person> people = new List<Person>
            {
                new Person { Id = 1, Name = "Alice",   BirthDate = new DateTime(1990, 5, 21) },
                new Person { Id = 2, Name = "Bob",     BirthDate = new DateTime(1985, 12, 3) },
                new Person { Id = 3, Name = "Charlie", BirthDate = new DateTime(2000, 1, 15) }
            };

            // Set import options to include header row
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true
            };

            // Import the custom objects into the worksheet starting at cell A1
            worksheet.Cells.ImportCustomObjects(people.ToArray(), 0, 0, importOptions);

            // Create a style with the desired custom number format for dates
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";

            // Apply the style to the entire BirthDate column (C column, index 2)
            // Use StyleFlag to limit the application to the number format only
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            worksheet.Cells.Columns[2].ApplyStyle(dateStyle, flag);

            // Save the workbook to a file
            workbook.Save("CustomObjectsWithFormattedDates.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Custom class with a DateTime property
    public class Person
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public DateTime BirthDate { get; set; }
    }
}
