// Title: How to hide columns G‑J and store a custom view in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells hide columns G to J and create a reusable custom view in C# | C# code to add a custom view that hides specific columns with Aspose.Cells | Using reflection to access the CustomViews collection in Aspose.Cells .NET | Save Excel file with hidden columns and a custom view using Aspose.Cells | Programmatically hide a range of columns and preserve view settings in Aspose.Cells
// Tags: hide columns range Aspose.Cells C# | add custom view Aspose.Cells .NET | custom view persistence Excel Aspose | reflection access CustomViews Aspose.Cells | save workbook with hidden columns Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example creates a new workbook, hides columns G through J on the first worksheet, attempts to add a custom view named "HiddenGtoJView" via reflection to the CustomViews collection, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide columns G (index 6) through J (index 9) – hide 4 columns starting at index 6
            sheet.Cells.HideColumns(6, 4);

            // Attempt to add a custom view if the API is available (using reflection)
            try
            {
                PropertyInfo customViewsProp = sheet.GetType().GetProperty("CustomViews");
                if (customViewsProp != null)
                {
                    object customViews = customViewsProp.GetValue(sheet);
                    MethodInfo addMethod = customViews?.GetType().GetMethod("Add", new[] { typeof(string) });
                    if (addMethod != null)
                    {
                        addMethod.Invoke(customViews, new object[] { "HiddenGtoJView" });
                        Console.WriteLine("Custom view added successfully.");
                    }
                    else
                    {
                        Console.WriteLine("Add method for custom views not found.");
                    }
                }
                else
                {
                    Console.WriteLine("CustomViews property not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                // Log any reflection-related errors
                Console.WriteLine($"Custom view not added: {ex.Message}");
            }

            // Define output file path
            string outputPath = "CustomViewWorkbook.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception e)
        {
            Console.WriteLine($"An error occurred: {e.Message}");
        }
    }
}
