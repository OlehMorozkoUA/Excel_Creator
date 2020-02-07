using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

namespace Excel_Create
{
    class PersonContext : DbContext
    {
        public DbSet<Person> People { get; set; }
    }
}
