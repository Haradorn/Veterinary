using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Veterinar
{
    public class Appointment
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public int? ClientId { get; set; }
        public Client Client{ get; set; }
        public int? PetId { get; set; }
        public Pet Pet{ get; set; }
        public int? VaccineId { get; set; }
        public Vaccine Vaccine{ get; set; }
        public string Service { get; set; }
        public string WhatHurt { get; set; }
        public string WhatWasDone { get; set; }
        public string WhatNeedToDo { get; set; }
        public override string ToString()
        {
            return Date.ToString() + " " + Client + " " + Pet;
        }
    }
}
