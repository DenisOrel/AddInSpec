using System.Collections.Generic;

namespace AddInSpec
{
    public class Assembly
    {
        public string Filename { get; set; }
        public List<AssemblyConfiguration> Configurations { get; set; }
    }
}