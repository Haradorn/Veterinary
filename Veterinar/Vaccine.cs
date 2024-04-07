using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Veterinar
{
    public class Vaccine
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public List<Appointment> Appointments { get; set; }
        public override string ToString()
        {
            return Name;
        }

    }
}
