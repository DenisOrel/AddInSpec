using System.Collections.Generic;

namespace AddInSpec
{
    internal class Assembly
    {
        public string Configuration { get; set; }
        public string Filename { get; set; }
        public Dictionary<string, string> Properties { get; set; }
        public List<Component> Components { get; set; }
    }
}
