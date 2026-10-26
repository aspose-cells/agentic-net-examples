// Title: How to delete an external connection by description from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# remove external connection with specific description using Aspose.Cells | Aspose.Cells delete external data connection from worksheet by description | How to access ExternalConnections collection via reflection in Aspose.Cells | Programmatically remove external connection from Excel file with Aspose.Cells .NET | Find and delete external connection named "My External Connection" in C# Aspose.Cells
// Tags: Aspose.Cells external connection removal API | C# reflection access ExternalConnectionCollection | programmatic deletion of Excel external connections | workbook external connections manipulation Aspose.Cells | saving workbook after external connection update

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads an Excel workbook, uses reflection to obtain the worksheet's ExternalConnections collection, searches for a connection whose Description matches a given value, removes that connection with the collection's RemoveAt method, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string targetDescription = "My External Connection";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: Workbook contains no worksheets.");
                return;
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Try to obtain the ExternalConnections property via reflection
            PropertyInfo externalConnProp = sheet.GetType().GetProperty("ExternalConnections", BindingFlags.Public | BindingFlags.Instance);
            if (externalConnProp == null)
            {
                Console.WriteLine("External connections are not supported in this version of Aspose.Cells.");
            }
            else
            {
                // Retrieve the collection (type is ExternalConnectionCollection in supported versions)
                object connectionsObj = externalConnProp.GetValue(sheet);
                if (connectionsObj == null)
                {
                    Console.WriteLine("No external connections collection found.");
                }
                else
                {
                    // Use reflection to work with the collection without compile‑time dependency
                    Type connCollectionType = connectionsObj.GetType();
                    PropertyInfo countProp = connCollectionType.GetProperty("Count");
                    MethodInfo getItemMethod = connCollectionType.GetMethod("get_Item");
                    MethodInfo removeAtMethod = connCollectionType.GetMethod("RemoveAt");

                    int count = (int)countProp.GetValue(connectionsObj);
                    int indexToRemove = -1;

                    for (int i = 0; i < count; i++)
                    {
                        object conn = getItemMethod.Invoke(connectionsObj, new object[] { i });
                        PropertyInfo descProp = conn.GetType().GetProperty("Description");
                        string description = descProp?.GetValue(conn) as string;

                        if (description == targetDescription)
                        {
                            indexToRemove = i;
                            break;
                        }
                    }

                    if (indexToRemove >= 0 && removeAtMethod != null)
                    {
                        removeAtMethod.Invoke(connectionsObj, new object[] { indexToRemove });
                        Console.WriteLine($"External connection \"{targetDescription}\" removed.");
                    }
                    else
                    {
                        Console.WriteLine($"External connection \"{targetDescription}\" not found.");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
