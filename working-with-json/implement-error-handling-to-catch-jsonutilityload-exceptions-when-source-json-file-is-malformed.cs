// Title: Add try‑catch error handling for malformed JSON when loading a Workbook with Aspose.Cells in C#
// AI Prompts: Generate C# code that uses Aspose.Cells JsonLoadOptions to load a JSON file into a Workbook and catches any JsonUtility.Load exceptions, logging the error message. | Show how to verify that a JSON file exists before invoking the Workbook constructor and return null if the file is missing or loading fails. | Demonstrate saving the Workbook to an .xlsx file only after a successful JSON load, with a separate try‑catch block for the Save method.
// Common Searches: aspocells c# catch JsonUtility.Load error when JSON is malformed | how to verify JSON file exists before loading into Aspose.Cells workbook | example of using JsonLoadOptions with try‑catch to handle bad JSON in Aspose.Cells
// Tags: manage JsonUtility.Load failures Aspose.Cells | malformed JSON handling with JsonLoadOptions C# | verify JSON file presence prior to workbook creation | try‑catch workbook save Aspose.Cells | load JSON into Workbook with robust error handling

using System;
using System.IO;
using Aspose.Cells;

// C# example that checks for a JSON file, loads it into an Aspose.Cells Workbook using JsonLoadOptions, catches malformed‑JSON and other load errors, and saves to Excel only when loading succeeds.
class JsonLoader
{
    // Loads a JSON file into a Workbook with error handling for malformed JSON.
    public static Workbook LoadJson(string jsonFilePath)
    {
        // Verify that the JSON file exists to avoid FileNotFoundException.
        if (!File.Exists(jsonFilePath))
        {
            Console.WriteLine($"File not found: {jsonFilePath}");
            return null;
        }

        try
        {
            // Use JsonLoadOptions to load JSON data.
            JsonLoadOptions loadOptions = new JsonLoadOptions();
            Workbook workbook = new Workbook(jsonFilePath, loadOptions);
            return workbook;
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading (e.g., malformed JSON).
            Console.WriteLine($"Error loading JSON: {ex.Message}");
            return null;
        }
    }

    // Example usage.
    static void Main()
    {
        string jsonPath = "data.json";

        Workbook wb = LoadJson(jsonPath);
        if (wb != null)
        {
            try
            {
                // Save the workbook to an Excel file.
                string outputPath = "output.xlsx";
                wb.Save(outputPath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("Workbook could not be created due to JSON errors.");
        }
    }
}
