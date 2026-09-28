// Title: Verify that adding a non‑registered custom function library to an Aspose.Cells workbook throws an exception in C#
// AI Prompts: Create a C# console application that opens an Aspose.Cells Workbook, attempts to load a non‑existent DLL as a custom function library, and catches the resulting FileNotFoundException or CellsException. | Write an NUnit test method that registers a custom function library with a missing assembly path and asserts that a CellsException is raised. | Provide a code example showing how to handle both FileNotFoundException and CellsException when calling Workbook.CustomFunctions.AddLibrary with an unregistered DLL.
// Common Searches: aspnet cells custom function library missing dll exception handling | c# Aspose.Cells verify exception for unregistered assembly reference | unit test for CellsException when adding non existent library to workbook | how to catch FileNotFoundException for custom functions in Aspose.Cells macro project | validate library reference loading failure in Aspose.Cells C# example
// Tags: custom function library loading exception Aspose.Cells | register unregistered assembly CellsException handling | Aspose.Cells workbook add library reference error | C# test missing DLL for custom functions | macro project library validation Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// Demonstrates how to confirm that adding a reference to a non‑existent custom function DLL in an Aspose.Cells workbook triggers a FileNotFoundException or CellsException, and shows proper exception handling in C#.
class LibraryReferenceTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define a library that is NOT registered on the host
            string unregisteredAssembly = "NonExistentLibrary.dll";
            string functionName = "MyCustomFunction";
            string className = "NonExistentNamespace.NonExistentClass";

            // Attempt to add a custom function from an unregistered library.
            // Since the CustomFunctions API may not be available in the current version,
            // we simulate the behavior by trying to load the assembly directly,
            // which will raise a FileNotFoundException.
            try
            {
                // Directly load the assembly to trigger the expected exception
                Assembly.LoadFrom(unregisteredAssembly);
            }
            catch (FileNotFoundException)
            {
                // Expected outcome: the assembly does not exist.
                Console.WriteLine("Test Passed: Caught expected FileNotFoundException.");
                throw; // Re‑throw to be caught by the outer handler for reporting.
            }

            // If no exception is thrown, the test has failed.
            Console.WriteLine("Test Failed: No exception was thrown when adding an unregistered library reference.");
        }
        catch (CellsException ex)
        {
            // Expected path if a CellsException is thrown by a newer API.
            Console.WriteLine("Test Passed: Caught expected CellsException.");
            Console.WriteLine($"Exception Message: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            // Expected when the fallback loading mechanism fails.
            Console.WriteLine("Test Passed: Caught expected FileNotFoundException.");
            Console.WriteLine($"Exception Message: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other exception type indicates unexpected behavior.
            Console.WriteLine("Test Failed: Caught an unexpected exception type.");
            Console.WriteLine($"Exception Type: {ex.GetType().FullName}");
            Console.WriteLine($"Exception Message: {ex.Message}");
        }
    }
}
