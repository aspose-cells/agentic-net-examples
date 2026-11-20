// Title: Export a list of Person objects with multiple addresses to an XLSX file by flattening nested collections into separate columns using Aspose.Cells for .NET
// Common Searches: c# export list of objects with nested address collections to Excel using Aspose.Cells | how to flatten a Person object with multiple addresses into separate worksheet columns in Aspose.Cells | generate dynamic Excel columns based on the maximum number of items in a collection with Aspose.Cells .NET | map custom C# objects to a flat dictionary for spreadsheet export with Aspose.Cells | create fixed address columns for varying address counts when writing to XLSX in Aspose.Cells
// Tags: flatten nested objects to Excel columns Aspose.Cells | dynamic column creation from collection size Aspose.Cells | export custom C# objects to XLSX Aspose.Cells | object‑to‑dictionary mapping for spreadsheet export Aspose.Cells | auto‑fit worksheet columns after data write Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsNestedImport
{
    // Sample custom classes
    // The example defines Person and Address classes, determines the maximum number of addresses across all persons, and uses a mapping function to convert each Person into a flat dictionary with columns for Name, Age, and a fixed set of AddressN_Street and AddressN_City fields. It writes a header row and data rows to the first worksheet of a new Workbook, auto‑fits the columns, and saves the file as PeopleFlattened.xlsx.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public List<Address> Addresses { get; set; }
    }

    public class Address
    {
        public string Street { get; set; }
        public string City { get; set; }
    }

    class Program
    {
        // Flattens a Person object into a dictionary of column names and values.
        static Dictionary<string, object> MapPersonToFlatColumns(Person person, int maxAddressCount)
        {
            // Guard against null Person
            if (person == null)
                person = new Person { Name = string.Empty, Age = 0, Addresses = new List<Address>() };

            var dict = new Dictionary<string, object>
            {
                { "Name", person.Name ?? string.Empty },
                { "Age", person.Age }
            };

            // Ensure a fixed number of address columns (maxAddressCount)
            for (int i = 0; i < maxAddressCount; i++)
            {
                if (i < (person.Addresses?.Count ?? 0) && person.Addresses[i] != null)
                {
                    dict[$"Address{i + 1}_Street"] = person.Addresses[i].Street ?? string.Empty;
                    dict[$"Address{i + 1}_City"]   = person.Addresses[i].City ?? string.Empty;
                }
                else
                {
                    // Fill empty cells for missing addresses
                    dict[$"Address{i + 1}_Street"] = string.Empty;
                    dict[$"Address{i + 1}_City"]   = string.Empty;
                }
            }

            return dict;
        }

        static void Main(string[] args)
        {
            try
            {
                // Sample data with nested collections
                var people = new List<Person>
                {
                    new Person
                    {
                        Name = "Alice",
                        Age = 30,
                        Addresses = new List<Address>
                        {
                            new Address { Street = "123 Main St", City = "Springfield" },
                            new Address { Street = "456 Oak Ave", City = "Shelbyville" }
                        }
                    },
                    new Person
                    {
                        Name = "Bob",
                        Age = 45,
                        Addresses = new List<Address>
                        {
                            new Address { Street = "789 Pine Rd", City = "Capital City" }
                        }
                    },
                    new Person
                    {
                        Name = "Charlie",
                        Age = 28,
                        Addresses = null // No addresses
                    }
                };

                // Determine the maximum number of addresses any person has (used to create consistent columns)
                int maxAddressCount = 0;
                foreach (var p in people)
                {
                    int count = p.Addresses?.Count ?? 0;
                    if (count > maxAddressCount) maxAddressCount = count;
                }

                // Create a new workbook (lifecycle rule)
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];
                sheet.Name = "People";

                // Prepare header row using a dummy person (no addresses needed)
                var headerDict = MapPersonToFlatColumns(
                    new Person { Name = "Name", Age = 0, Addresses = new List<Address>() },
                    maxAddressCount);

                int colIndex = 0;
                foreach (var header in headerDict.Keys)
                {
                    sheet.Cells[0, colIndex].PutValue(header);
                    colIndex++;
                }

                // Write data rows
                int rowIndex = 1;
                foreach (var person in people)
                {
                    var rowDict = MapPersonToFlatColumns(person, maxAddressCount);
                    colIndex = 0;
                    foreach (var value in rowDict.Values)
                    {
                        sheet.Cells[rowIndex, colIndex].PutValue(value);
                        colIndex++;
                    }
                    rowIndex++;
                }

                // Auto-fit columns for better appearance
                sheet.AutoFitColumns();

                // Save the workbook to a file (lifecycle rule)
                workbook.Save("PeopleFlattened.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
