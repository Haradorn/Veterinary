using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Veterinar
{
    public class Pet
    {
        public int Id { get; set; }
        public string PetName { get; set; }
        public string Breed { get; set; }
        public DateTime Date { get; set; }
        public int? ClientId { get; set; }
        public Client Client { get; set; }
        public override string ToString()
        {
            return PetName;
        }
    }
}
