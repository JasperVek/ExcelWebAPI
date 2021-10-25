using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWebAPI.Domain
{
    public class AllInvestments : IModel
    {
        public AllInvestments()
        {
            PersonalInvestment = new PersonalInvestment();
            NonPersonalInvestment = new NonPersonalInvestment();
        }
        public PersonalInvestment PersonalInvestment { get; set; }
        public NonPersonalInvestment NonPersonalInvestment { get; set;  }

        public List<IItem> ItemList { get; set; }
    }
}
