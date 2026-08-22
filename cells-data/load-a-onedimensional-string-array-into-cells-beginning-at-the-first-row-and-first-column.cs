// Title: Import a one‑dimensional string array into the first row of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a string[] into the first row of a worksheet with Cells.ImportObjectArray and saves the workbook. | Create a snippet that converts a string array to an object array and imports it horizontally starting at cell A1 using Aspose.Cells. | Write error‑handling logic for importing a one‑dimensional array into an Excel sheet with Aspose.Cells and logging any failures.
// Common Searches: Aspose.Cells C# import string array into first row of worksheet | How to write a string[] to Excel cells starting at A1 with Aspose.Cells | Using Cells.ImportObjectArray to place data horizontally in a sheet | Convert string array to object array for Aspose.Cells import example | Save workbook after importing a one‑dimensional array with Aspose.Cells
// Tags: import string array Aspose.Cells | ImportObjectArray horizontal import C# | convert string[] to object[] Aspose.Cells | load one-dimensional array into Excel cells | save workbook after array import Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, converts a string[] to an object[] and uses Cells.ImportObjectArray to write the values horizontally starting at cell A1, then saves the file as OneDimensionalStringArray.xlsx.
    public class LoadOneDimensionalStringArray
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // One‑dimensional string array to import
                string[] stringArray = new string[] { "Alpha", "Beta", "Gamma", "Delta" };

                // Convert to object[] because ImportObjectArray expects an object array
                object[] data = Array.ConvertAll(stringArray, s => (object)s);

                // Import the array starting at the first row (0) and first column (0) horizontally
                // isVertical = false means data will be placed across columns in the first row
                cells.ImportObjectArray(data, 0, 0, false);

                // Save the workbook to a file
                workbook.Save("OneDimensionalStringArray.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadOneDimensionalStringArray.Run();
        }
    }
}
