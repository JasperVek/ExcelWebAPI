using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWebAPI.Domain
{
    public class Investment : IItem
    {
        public Investment()
        {
        }
        public string InvestmentType { get; private set; }
        public string Symbol { get; private set; }
        public string CurrentPercentOfPortfolio { get; private set; }
        public string TargetPercentOfPortfolio { get; private set; }
        public string CurrentVsTarget { get; private set; }

        public void FillItem(List<Tuple<string, string>> values)
        {
            foreach (var item in values)
            {
                switch (item.Item1)
                {
                    case nameof(InvestmentType):
                        InvestmentType = item.Item2;
                        break;
                    case nameof(Symbol):
                        Symbol = item.Item2;
                        break;
                    case nameof(CurrentPercentOfPortfolio):
                        CurrentPercentOfPortfolio = item.Item2;
                        break;
                    case nameof(TargetPercentOfPortfolio):
                        TargetPercentOfPortfolio = item.Item2;
                        break;
                    case nameof(CurrentVsTarget):
                        CurrentVsTarget = item.Item2;
                        break;

                    default:
                        break;
                }
            }
        }
    }
}
