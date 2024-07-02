using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NWTARules
{
    public class Function
    {
        public string Name { get; set; }
        public string ID { get; set; }

        public List<NameAndRole> Staff {  get; set; } = new List<NameAndRole>();

        public Function(string name, string id) 
        {
            Name = name;
            ID = id;   
        }
    }

    public class NameAndRole
    {
        public string Name { get; set; }
        public string Role { get; set; }
        public bool NewMember { get; set; } = false;

        public NameAndRole(string name, string role)
        {
            Name = name; 
            Role = role;
        }
    }
}
