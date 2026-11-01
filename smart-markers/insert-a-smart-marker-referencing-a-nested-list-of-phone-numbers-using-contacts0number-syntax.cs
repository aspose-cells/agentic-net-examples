// Title: How to insert a smart marker that references the first contact's phone number (contacts[0].number) in an Aspose.Cells workbook using C#
// Common Searches: C# Aspose.Cells insert smart marker for contacts[0].number array element | How to reference a nested list value in Aspose.Cells smart marker | Aspose.Cells smart marker syntax for first item in a collection in .NET | Insert smart marker into cell A1 to pull phone number from contacts list using Aspose.Cells | Using Aspose.Cells smart markers with nested objects in C# workbook
// Tags: insert smart marker Aspose.Cells C# | smart marker array element syntax | Aspose.Cells nested list binding | C# workbook cell smart marker | contacts[0].number marker example

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Aspose.Cells workbook, adds a smart marker '&=contacts[0].number' to cell A1 to pull the first contact's phone number from a nested list, and saves the file as SmartMarkerOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = CreateWorkbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that references the first contact's phone number
            // Syntax: contacts[0].number
            sheet.Cells["A1"].PutValue("&=contacts[0].number");

            // Save the workbook
            SaveWorkbook(workbook, "SmartMarkerOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a new workbook (empty or from a template if needed)
    private static Workbook CreateWorkbook()
    {
        // No template file is required; just instantiate a new workbook
        return new Workbook();
    }

    // Saves the workbook to the specified path, ensuring the directory exists
    private static void SaveWorkbook(Workbook workbook, string outputPath)
    {
        try
        {
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
            throw;
        }
    }
}
