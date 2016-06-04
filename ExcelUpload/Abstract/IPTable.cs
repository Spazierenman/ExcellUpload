using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUpload.Abstract
{
    public interface IPTable
    {
        string Name { get;}

        void AddHeader(string Name);

        string Flow { set; }

        object GetDataList { get; }

    }
}
