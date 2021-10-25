using ExcelWebAPI.Domain;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelWebAPI
{
    public class DemoReader : IExcelReader
    {
        // fill a IModel with recursion
        public IModel ReadModel(IModel model, ExcelWorksheet worksheet)
        {
            // test if the Model contains other Models, or Items
            List <IModel> listSubModels = ExtractSubModels(model);

            // recursion
            if(listSubModels != null && listSubModels.Count > 0)
            {
                foreach (var sub in listSubModels)
                {
                     sub.ItemList = ReadModel(sub, worksheet).ItemList;
                }
            }

            // base: IModel has a list with IItems, which returns this IItems list after having found the items in Excel
             return FillModel(model, worksheet);
        }

        // finds and extracts the properties that are IModel
        private List<IModel> ExtractSubModels(IModel model)
        {
            List<IModel> modelList = new List<IModel>();

            foreach (var property in model.GetType().GetProperties())
            {
                // TODO: hopen dat de Models eerst zijn gedeclareerd..
                if(typeof(IModel).IsAssignableFrom(property.PropertyType))
                {
                    if (typeof(List<IItem>).IsAssignableFrom(property.PropertyType) == false)
                    {
                        modelList.Add(property.GetValue(model) as IModel);
                    }
                }
            }

            return modelList;
        }

        // fill model with the List<IItem>
        public IModel FillModel(IModel model, ExcelWorksheet excelFile)
        {

            IItem item = GetListItemType(model);

            if (item != null)
            {
                model.ItemList = ReadItemsFromExcel(item, excelFile);
            }
            return model;
        }

        // this reader reads all cells under properties, and searches down until empty or the same properties occur
        private List<IItem> ReadItemsFromExcel(IItem item, ExcelWorksheet excelFile)
        {
            List<IItem> endList = new List<IItem>();

            List<string> propNames = GetPropNamesAsList(item);

            // name, row, column
            List<Tuple<string, int, int>> propPositionList = GetPropPositions(propNames, excelFile);

            var rowCount = propPositionList[0].Item2 + 1;

            bool reading = true;
            
            // if FillItem is null: stop
            while (reading)
            {
                var itemToAdd = Activator.CreateInstance(item.GetType());
                itemToAdd = FillItem(itemToAdd as IItem, propPositionList, rowCount, excelFile);
                if(itemToAdd == null)
                {
                    break;
                }
                endList.Add(itemToAdd as IItem);

                rowCount++;
            }
            return endList;
        }

        // extract the property names from Excel, based on IItem that is provided
        private List<string> GetPropNamesAsList(IItem item)
        {
            List<string> namesList = new List<string>();
            foreach (var property in item.GetType().GetProperties())
            {
                namesList.Add(property.Name);
            }

            return namesList;
        }

        // Locate the property positions in Excel
        private List<Tuple<string, int, int>> GetPropPositions(List<string> propNames, ExcelWorksheet worksheet)
        {
            List<Tuple<string, int, int>> propList = new List<Tuple<string, int, int>>();

            foreach (var item in propNames)
            {
                // zoek de cell waar de value van de cell overeen komt met de item
                var cellPositions = SearchCellByValue(item, worksheet, 20);

                propList.Add(Tuple.Create(item, cellPositions.Item1, cellPositions.Item2));
            }
            return propList;
        }

        // search and fill next value in Excel to fill the IItems list
        private IItem FillItem(IItem item, List<Tuple<string, int, int>> propPositionList, int rowCount, ExcelWorksheet worksheet)
        {
           // row, column
            List<Tuple<string, string>> itemValuesList = new List<Tuple<string, string>>();

            foreach (var propertyInfo in propPositionList)
            {
                var valueFromCell = worksheet.Cells[rowCount, propertyInfo.Item3].Value;

                SearchNewValueForItemList(itemValuesList, propertyInfo, valueFromCell);
            }

            if (itemValuesList.Count < propPositionList.Count)
            {
                return null;
            }

            // object knows how to assign its values based on IItem interface
            item.FillItem(itemValuesList);
            return item;
        }

        // Search the next Item Value
        private static void SearchNewValueForItemList(List<Tuple<string, string>> itemValuesList, Tuple<string, int, int> propertyInfo, object valueFromCell)
        {
            // don't add if cell is null, or matches a property name
            if (valueFromCell != null && valueFromCell.ToString() != "")
            {
                string valueToCompare = String.Concat(valueFromCell.ToString().Where(c => !Char.IsWhiteSpace(c)));
                if (propertyInfo.Item1.ToLower() != valueFromCell.ToString().ToLower())
                {
                    // name + value van de Cell
                    Tuple<string, string> itemContent = Tuple.Create(propertyInfo.Item1, valueFromCell.ToString());
                    itemValuesList.Add(itemContent);
                }
            }
        }

        // search method for string value in Excel, with a cap on rows for the writer to decide, and a hardcoded cap for columns (100)
        private Tuple<int,int> SearchCellByValue(string toSearch, ExcelWorksheet worksheet, int rowRange)
        {
            int column = 1;
            int row = 1;
            bool found = false;

            // TODO max row en max column als echt limiet
            // voorlopig hardcoded op 100...
            while(found == false && column < 100)
            {
                object value;
                try
                {
                    value = worksheet.Cells[row, column].Value;
                }
                catch (Exception e)
                {
                    throw;
                }
                if (value != null)
                {
                    // lower en verwijderen whitespaces..
                    string valueToCompare = String.Concat(value.ToString().Where(c => !Char.IsWhiteSpace(c)));
                    
                    if(valueToCompare.ToLower() == toSearch.ToLower())
                    {
                        return Tuple.Create(row, column);
                    }
                }
                row++;

                if (row > rowRange)
                {
                    row = 1;
                    column++;
                }
            }

            return null;
        }

        // extract which concrete IItem this is
        private IItem GetListItemType(IModel model)
        {
            foreach (var property in model.GetType().GetProperties())
            {
                if (typeof(List<IItem>).IsAssignableFrom(property.PropertyType))
                {
                    List<IItem> listItems = property.GetValue(model) as List<IItem>;
                    if(listItems?.Count > 0 && typeof(IItem).IsAssignableFrom(listItems[0].GetType()))
                    {
                        return listItems[0];
                    }
                }
            }
            return null;
        }
    }
}
