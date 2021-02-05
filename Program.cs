using System;
using System.Collections.Generic;

namespace ConsumeTypes
{
    class Program
    {
        static void Main(string[] args)
        {
            //Boxing and Unboxing Method:
            BoxingAndUnboxing();

            //Conversions and casting
            Conversions();

            //Using Dynamic Types: Office Automation API's
            Console.WriteLine("Do you want to see how the Office Automation API for excel is used? Go check out the code in line 91(DisplayInExcel method).");
            //Create Excel object(records)
            //Add the following code to add reords to excel worksheet
            // var entities = new List<dynamic>
            //         {
            //                     new
            //                     {
            //                         ColumnA = 1,
            //                         ColumnB = "Foo"
            //                     },
            //                     new
            //                     {
            //                         ColumnA= 2,
            //                         ColumnB= "Bar"
            //                     }
            //         };

            // //Pass created object as parameter
            // DisplayInExcel(entities);
        }

        //Boxing and Unboxing
        //Boxing: Process of taking value type, placing it in new object on heap, storing ref to it on stack
        //Unboxing: Takes item from heap, returns value that contains value from heap
        public static void BoxingAndUnboxing()
        {
            //Boxing int value:
            int i = 42;
            object o = i;
            Console.WriteLine("Boxed: {0}", o);

            //Unboxing int value
            int x = (int)o;
            Console.WriteLine("Unboxed: {0}", x);
        }

        //Converting between different types
        //4 Ways of conversion: 
        //1) Implicit conversions, 2) explicit conversions, 3) User-Defined conversions and 4) Conversion with helper class
        public static void Conversions()
        {
            // 1) Implicit conversions: When conversion is legal, allowed
            int i = 42;
            double d = i;

            Console.WriteLine("Implicit conversion(int to double): {0}", d);

            // 2) Explicit conversions: when conversion is not allowed, needs to be casted
            double d2 = 42.7;
            int i2 = (int)d2;
            
            Console.WriteLine("Explicit conversion(double to int): {0}", i2);

            // 3) User-Defined conversions: When working with own types, you use both Impplicit- and Explicit conversions
            var money = new Money(12.99M);
            money.Amount = 12.99M;

            // 4) Conversions with Helper class(BitConverter- and Converter class from System namespace)

            //Convert class
            int value1 = Convert.ToInt32("42");
            Console.WriteLine("Convert.ToInt32 value: {0}", value1);

            //Parse class
            int value2 = int.Parse("42");
            Console.WriteLine("int.Parse value: {0}", value2);

            //TryParse(returns value and true/false)
            bool success = int.TryParse("42", out int value3);
            Console.WriteLine("Convertion succesfull: {0}, Value: {1}", success, value3);


        }

        //Dynamic Types: Excel worksheet
        static void DisplayInExcel(IEnumerable<dynamic> entities)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Add();

            dynamic workSheet = excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "Header A";
            workSheet.Cells[1, "B"] = "Header B";

            var row = 1;
            foreach (var entity in entities)
            {
                row++;
                workSheet.Cells[row, "A"] = entity.ColumnA;
                workSheet.Cells[row, "B"] = entity.ColumnB;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
        }


        class Money
        {
            //Amount
            public decimal Amount { get; set; }

            //Parameterized constructor
            public Money(decimal amount)
            {
                Amount = amount;
            }

            //Implicit operator
            public static implicit operator decimal(Money money)
            {
                Console.WriteLine("Implicit operator: {0}", money.Amount);
                return money.Amount;
            }

            //Explicit operator
            public static explicit operator int(Money money)
            {
                Console.WriteLine("Explicit operator: {0}", (int)money.Amount);
                return (int)money.Amount;
            }
        }  
    }
}
