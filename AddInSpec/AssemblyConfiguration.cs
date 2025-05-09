using System.Collections.Generic;

namespace AddInSpec
{
    public class AssemblyConfiguration
    {
        public string Configuration { get; set; }
        public Dictionary<string, string> Properties { get; set; }
        public List<Component> Components { get; set; }
    }
}