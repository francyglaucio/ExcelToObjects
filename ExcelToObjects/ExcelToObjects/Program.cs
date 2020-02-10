using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToObjects
{
    class Program
    {
        static void Main(string[] args)
        {
            String fileName = @"C:\projetos\Products.xlsx";
            List<ProductModel> products = new List<ProductModel>();
            products = ExcellToObjects.ConvertToObject<ProductModel>(fileName);
           
        }
    }
}
