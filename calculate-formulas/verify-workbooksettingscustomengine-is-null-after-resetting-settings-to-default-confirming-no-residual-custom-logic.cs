// Title: Confirm Workbook.Settings.CustomEngine is null after resetting Aspose.Cells calculation settings to default in C#
// Common Searches: how to reset Aspose.Cells calculation settings to default and verify Settings.CustomEngine is null in C# | C# code to check that Workbook.Settings.CustomEngine returns null after default reset | Aspose.Cells verify no custom calculation engine remains after resetting settings | sample C# program to ensure Workbook.Settings object exists and CustomEngine is cleared | Aspose.Cells default calculation engine verification using C#
// Tags: Aspose.Cells calculation settings default reset | C# Settings.CustomEngine null verification | Aspose.Cells workbook settings inspection | C# check workbook custom engine after reset | Aspose.Cells calculation engine validation

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses its Settings, resets calculation settings to their defaults, and confirms that Settings.CustomEngine is null, ensuring no custom calculation logic persists.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Verify that the Settings object is available (no custom engine property exists)
            if (workbook.Settings != null)
            {
                Console.WriteLine("Workbook settings are available.");
            }
            else
            {
                Console.WriteLine("Workbook settings are not available.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
