using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_Task
{
    public class TrasmittalFile
    {
        public string Dir { get; set; }
        public string Name { get; set; }
        public string Ext { get; set; }


        public TrasmittalFile(string dir, string name, string ext)
        {
            Dir = dir;
            Name = name;
            Ext = ext;
        }
    }
}
