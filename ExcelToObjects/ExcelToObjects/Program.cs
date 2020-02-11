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
            String fileName = @"..\..\files\Products2.xlsx";
            List<ProductModel> products = new List<ProductModel>();
            products = new ExcellToObjects(fileName).ConvertToObject<ProductModel>();

            foreach (var obj in products)
            {
                Console.WriteLine(obj.Id + " " + obj.Name + " " + obj.Price + " " + obj.Quantity + " " + obj.Teste);
            }
            Console.ReadKey();

        }
    }
}
