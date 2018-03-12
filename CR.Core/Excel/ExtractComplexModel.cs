using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CR.Core.Excel
{
    public class ExtracterComplexModel
    {
        public ExtracterComplexModel()
        {
            Messages = new List<string>();
        }

        public Type ModelType { get; set; }

        public List<string> Messages { get; set; }

        public object Data { get; set; }
    }
}
