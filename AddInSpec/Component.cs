﻿using System.Collections.Generic;

namespace AddInSpec
{
    public class Component
    {
        public string Configuration { get; set; }
        public string Filename { get; set; }
        public string Quantity { get; set; }
        public Dictionary<string, string> Properties { get; set; }
    }
}
