using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CR.Core.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute(string field)
        {
            this.Field = field;
        }

        public ExcelColumnAttribute(string field, object def, Type deftype = null)
        {
            this.Field = field;
            this.Default = def;
            this.DefaultType = deftype;
        }

        public string Field { get; set; }

        public object Default { get; set; }

        public Type DefaultType { get; set; }

        private bool _isrequire = false;
        public bool IsRequire { get { return _isrequire; } set { _isrequire = value; } }
    }
}
