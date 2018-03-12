using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CR.Core.Excel
{
    public interface IWriteHandler
    {
        Workbook ExecuteWork(Workbook tp, Dictionary<string, object> param);
    }
}
