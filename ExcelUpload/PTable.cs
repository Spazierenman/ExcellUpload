using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUpload.Abstract;
using System.Reflection;
using System.ComponentModel.DataAnnotations;
using System.Globalization;

namespace ExcelUpload
{
    /// <summary>
    /// Класс, принимающий тип модели. Значение свойства 'Name' эквивалентно значению свойства WorkSheet 'Name'
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class PTable <T> : IPTable

        where T : new()

    {
        public PTable()
        {
            ModelType = typeof(T);
            Name = ModelType.Name;
            Columns = new List<PColumn>();
            Rows = new List<T>();
        }

        private Type ModelType { get; set; }

        private List<PColumn> Columns;

        private List<T> Rows;

        public string Name { get; set; }

        public void AddHeader(string Name)
        {
            PropertyInfo pi = ModelType.GetProperties().FirstOrDefault(s =>
            {

                DisplayAttribute dattr = (DisplayAttribute)Attribute.GetCustomAttribute(s, typeof(DisplayAttribute));

                return (s.Name == Name) || ((dattr != null) && (dattr.Name == Name));

            });



            if (pi != null)
            {
                MethodInfo mf = ModelType.GetMethods().FirstOrDefault(s => s.Name == pi.Name + "_");

                    Columns.Add(new PColumn()
                    {
                        Header = Name,
                        proinf = pi,
                        NeedConverter = (mf != default(MethodInfo)),
                        Converter = mf,
                        Order = Columns.Count
                    });
            }
        }


        private PColumn CurrentColumn { get; set; }

        private T CurrentRow { get; set; }

        private void FlowIncrement()
        {
            if ((CurrentColumn == null) || (CurrentColumn.Order >= Columns.Count - 1 ))
            {
                CurrentColumn = Columns.Single(s=> s.Order == 0);
                CurrentRow = new T();
                Rows.Add(CurrentRow);
            }
            else CurrentColumn = Columns.Single(s => s.Order == CurrentColumn.Order + 1);
        }

        public string Flow
        {
            set
            {
                FlowIncrement();

                object typevalue = value;

                if (CurrentColumn.NeedConverter)
                {
                    typevalue =  CurrentColumn.Converter.Invoke(CurrentRow, new object[] { typevalue });
                }

                if (CurrentColumn.proinf.PropertyType == typeof(DateTime))
                {
                    CurrentColumn.proinf.SetValue(CurrentRow, ParseDateTime(typevalue.ToString()), null);
                }
                else
                    CurrentColumn.proinf.SetValue(CurrentRow, Convert.ChangeType(typevalue, CurrentColumn.proinf.PropertyType), null);

            }
        }

        private static DateTime ParseDateTime(string cellValue)
        {
            DateTime date = default(DateTime);

            double value = default(double);

            if (double.TryParse(cellValue, out value))
            {
                date = DateTime.FromOADate(value);
            }
            else
            {
                DateTime.TryParse(cellValue, out date);
            }

            return date;
        }

        public object GetDataList
        {
            get
            {
                return Rows;
            }
        }

        }
}
