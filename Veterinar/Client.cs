using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Veterinar
{
    public class Client
    {
        public int Id { get; set; } 
        public string Name { get; set; }
        public string Address { get; set; } 
        public string Phone { get; set; }
        public List<Pet> Pets { get; set; }
        public override string ToString()
        {
            return Name;
        }
    }
}
