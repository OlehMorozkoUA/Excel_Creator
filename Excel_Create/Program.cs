using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Create
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel excel = new Excel("d:\\Test_Excel.xlsx", "Test");
            PersonContext db = new PersonContext();
            excel.ExcelCreateTableBody(db.People.ToList());
            //for (int i = 0; i < 100; i++)
            //{
            //    db.People.Add(new Person { FirstName = "FirstName" + i.ToString(), LastName = "LastName" + i.ToString(), Age = 24 + i });
            //}
            //db.SaveChanges();
            foreach(Person person in db.People)
            {
                Console.WriteLine(person.FirstName);
            }
        }
    }
}
