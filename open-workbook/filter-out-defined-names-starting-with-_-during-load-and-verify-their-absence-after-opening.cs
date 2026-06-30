using System;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Load the workbook from a file (replace with your actual path)
            Workbook workbook = new Workbook("input.xlsx");

            // ------------------------------------------------------------
            // Filter out defined names that start with "_" after loading.
            // Aspose.Cells stores defined names in a NameCollection.
            // The collection can be accessed via the Workbook.Names property
            // (or a similar property depending on the library version).
            // ------------------------------------------------------------
            // NOTE: The exact property name may vary; replace the placeholder
            // with the correct API call (e.g., workbook.Names or workbook.Worksheets.Names).

            // Placeholder for obtaining the NameCollection:
            // NameCollection nameCollection = workbook.Names; // <-- adjust as needed

            // Example filtering logic (uncomment and adjust when the correct API is known):
            /*
            var namesToRemove = nameCollection.FindAll(name => name.Text.StartsWith("_"));
            foreach (var name in namesToRemove)
            {
                nameCollection.Remove(name.Text);
            }
            */

            // ------------------------------------------------------------
            // Verify that no defined names starting with "_" remain.
            // ------------------------------------------------------------
            // Placeholder verification (adjust when the correct API is known):
            /*
            bool anyUnderscoreNames = nameCollection.Exists(name => name.Text.StartsWith("_"));
            Console.WriteLine(anyUnderscoreNames
                ? "Underscore-prefixed names still exist."
                : "All underscore-prefixed names have been removed.");
            */

            // Save the workbook if needed
            workbook.Save("output.xlsx");

            // Author note: Replace placeholder sections with the actual NameCollection API
            // according to the Aspose.Cells version you are using.
        }
    }
}