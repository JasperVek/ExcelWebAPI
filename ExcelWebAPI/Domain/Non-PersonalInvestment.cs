using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWebAPI.Domain
{
    public class NonPersonalInvestment : IModel
    {
        public NonPersonalInvestment()
        {
            ItemList.Add(new Investment());
        }

        public List<IItem> ItemList { get; set; } = new List<IItem>();
    }
}
