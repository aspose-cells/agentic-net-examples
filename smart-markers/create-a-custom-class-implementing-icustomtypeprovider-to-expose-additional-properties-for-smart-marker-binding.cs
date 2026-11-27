// Title: Implement a custom ICustomTypeDescriptor to expose calculated fields for Aspose.Cells smart markers in C#
// Common Searches: how to use ICustomTypeDescriptor with Aspose.Cells WorkbookDesigner to add calculated columns | c# smart markers custom type provider example for adding IsAdult property | expose additional properties for smart marker binding using PropertyDescriptor in Aspose.Cells | aspnet cells smart markers custom data source with calculated fields | add DisplayName calculated property to smart marker data source in C#
// Tags: custom ICustomTypeDescriptor Aspose.Cells smart markers | calculated property IsAdult smart marker C# | WorkbookDesigner custom data source PropertyDescriptor | add DisplayName property via ICustomTypeDescriptor | smart marker binding custom type provider

using System;
using System.Collections.Generic;
using System.ComponentModel;
using Aspose.Cells;

namespace SmartMarkerCustomTypeProviderDemo
{
    // Simple POCO class that will be used as the original data source
    // The sample defines a Person POCO, a CalculatedPropertyDescriptor, and a PersonCustomProvider that implements ICustomTypeDescriptor to add read‑only IsAdult and DisplayName properties. A list of these providers is bound to Aspose.Cells smart markers through WorkbookDesigner, the markers are processed, and the populated workbook is saved as an XLSX file.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Person() { }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    // Custom PropertyDescriptor that exposes a calculated property (IsAdult)
    public class CalculatedPropertyDescriptor : PropertyDescriptor
    {
        private readonly Func<object, object> _getter;
        private readonly Action<object, object>? _setter;
        private readonly Type _propertyType;

        public CalculatedPropertyDescriptor(string name, Type propertyType,
            Func<object, object> getter, Action<object, object>? setter = null)
            : base(name, null)
        {
            _propertyType = propertyType;
            _getter = getter ?? throw new ArgumentNullException(nameof(getter));
            _setter = setter;
        }

        public override Type ComponentType => typeof(object);
        public override bool IsReadOnly => _setter == null;
        public override Type PropertyType => _propertyType;

        public override bool CanResetValue(object component) => false;
        public override object GetValue(object component) => _getter(component);
        public override void ResetValue(object component) { }
        public override void SetValue(object component, object value)
        {
            _setter?.Invoke(component, value);
        }
        public override bool ShouldSerializeValue(object component) => false;
    }

    // Wrapper that implements ICustomTypeDescriptor to expose additional properties
    public class PersonCustomProvider : ICustomTypeDescriptor
    {
        private readonly Person _person;

        public PersonCustomProvider(Person person)
        {
            _person = person ?? throw new ArgumentNullException(nameof(person));
        }

        // ICustomTypeDescriptor implementation – delegate most calls to the default descriptor
        public AttributeCollection GetAttributes() => TypeDescriptor.GetAttributes(_person, true);
        public string? GetClassName() => TypeDescriptor.GetClassName(_person, true);
        public string? GetComponentName() => TypeDescriptor.GetComponentName(_person, true);
        public TypeConverter? GetConverter() => TypeDescriptor.GetConverter(_person, true);
        public EventDescriptor? GetDefaultEvent() => TypeDescriptor.GetDefaultEvent(_person, true);
        public PropertyDescriptor? GetDefaultProperty() => TypeDescriptor.GetDefaultProperty(_person, true);
        public object? GetEditor(Type editorBaseType) => TypeDescriptor.GetEditor(_person, editorBaseType, true);
        public EventDescriptorCollection GetEvents(Attribute[]? attributes) => TypeDescriptor.GetEvents(_person, attributes, true);
        public EventDescriptorCollection GetEvents() => TypeDescriptor.GetEvents(_person, true);
        public PropertyDescriptorCollection GetProperties(Attribute[]? attributes) => GetCustomProperties(attributes);
        public PropertyDescriptorCollection GetProperties() => GetCustomProperties(null);
        public object GetPropertyOwner(PropertyDescriptor pd) => _person;

        // Build a PropertyDescriptorCollection that contains the original properties plus the calculated ones
        private PropertyDescriptorCollection GetCustomProperties(Attribute[]? attributes)
        {
            // Get the default properties of the underlying Person object
            PropertyDescriptorCollection baseProps = TypeDescriptor.GetProperties(_person, attributes, true);

            // Create a list to hold all descriptors
            List<PropertyDescriptor> allProps = new List<PropertyDescriptor>();
            foreach (PropertyDescriptor pd in baseProps)
            {
                allProps.Add(pd);
            }

            // Add a calculated read‑only property "IsAdult" (true if Age >= 18)
            var isAdultProp = new CalculatedPropertyDescriptor(
                "IsAdult",
                typeof(bool),
                component => ((Person)component).Age >= 18);

            allProps.Add(isAdultProp);

            // Add another calculated property "DisplayName" (e.g., "John (30)")
            var displayNameProp = new CalculatedPropertyDescriptor(
                "DisplayName",
                typeof(string),
                component =>
                {
                    var p = (Person)component;
                    return $"{p.Name} ({p.Age})";
                });

            allProps.Add(displayNameProp);

            return new PropertyDescriptorCollection(allProps.ToArray(), true);
        }
    }

    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Prepare sample data – a list of Person objects wrapped by
            //    PersonCustomProvider so that Aspose.Cells can see the extra
            //    properties via ICustomTypeDescriptor.
            // ------------------------------------------------------------
            List<PersonCustomProvider> data = new List<PersonCustomProvider>
            {
                new PersonCustomProvider(new Person("John Doe", 28)),
                new PersonCustomProvider(new Person("Jane Smith", 16)),
                new PersonCustomProvider(new Person("Bob Johnson", 45))
            };

            // ------------------------------------------------------------
            // 2. Create a workbook that contains smart markers.
            //    The markers reference both original and calculated properties.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Is Adult?");
            cells["D1"].PutValue("Display Name");

            // Smart marker rows – note the use of &=$ to indicate a smart marker
            cells["A2"].PutValue("&=$Name");
            cells["B2"].PutValue("&=$Age");
            cells["C2"].PutValue("&=$IsAdult");
            cells["D2"].PutValue("&=$DisplayName");

            // ------------------------------------------------------------
            // 3. Configure WorkbookDesigner, set the data source and process.
            // ------------------------------------------------------------
            try
            {
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                // The name "People" will be used in the smart marker expression.
                designer.SetDataSource("People", data);
                // Process the entire workbook (preserve unrecognized markers = false)
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during smart marker processing: {ex.Message}");
                return;
            }

            // ------------------------------------------------------------
            // 4. Save the result and output a simple console confirmation.
            // ------------------------------------------------------------
            string outputPath = "SmartMarkerCustomProviderResult.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }

            // ------------------------------------------------------------
            // 5. (Optional) Display the generated values in the console for verification.
            // ------------------------------------------------------------
            Console.WriteLine("\nGenerated Data:");
            for (int row = 1; row <= data.Count; row++)
            {
                string name = cells[$"A{row + 1}"].StringValue;
                string age = cells[$"B{row + 1}"].StringValue;
                string isAdult = cells[$"C{row + 1}"].StringValue;
                string displayName = cells[$"D{row + 1}"].StringValue;
                Console.WriteLine($"Row {row}: Name={name}, Age={age}, IsAdult={isAdult}, DisplayName={displayName}");
            }
        }
    }
}
