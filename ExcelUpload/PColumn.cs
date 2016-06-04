using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using ExcelUpload.Abstract;
using System.Collections;

namespace ExcelUpload
{
    class PColumn
    {


        public PropertyInfo proinf { get; set; }

        public int Order { get; set; }

        public string Header { get; set;}

        public bool NeedConverter { get; set; }

        public MethodInfo Converter { get; set; }


    }
}
