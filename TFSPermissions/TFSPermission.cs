using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFSPermissions
{
    class Project
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public List<Team> TeamNames { get; set; }
        
    }

    public class Team
    {
        //public int Id { get; set; }
        public string name { get; set; }
        public List<Member> Members { get; set; }
    }

    public class Member
    {
        //public int Id { get; set; }
        public string name { get; set; }
        public List<Group> Permissions { get; set; }
    }

    public class Group
    {
        //public int Id { get; set; }
        public string Permission { get; set; }
    }
}
