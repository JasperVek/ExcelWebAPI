using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWebAPI.Domain
{
    public interface IItem
    {
        void FillItem(List<Tuple<string, string>> values);
    }
}
