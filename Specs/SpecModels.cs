using System.Collections.Generic;

namespace LogiK3D.Specs
{
    public class PipingSpec
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Material { get; set; }
        public List<SpecComponent> Components { get; set; } = new List<SpecComponent>();

        public override string ToString()
        {
            return Name;
        }
    }

    public class SpecComponent
    {
        public string Type { get; set; } // PIPE, ELBOW, TEE, FLANGE, etc.
        public string DN { get; set; } // ex: "100"
        public double OuterDiameter { get; set; } // ex: 114.3
        public double Thickness { get; set; } // ex: 3.2
        public string Schedule { get; set; } // ex: "Sch 40"
        public string SAPCode { get; set; } // Code article
    }
}
