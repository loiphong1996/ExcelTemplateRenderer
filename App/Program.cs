using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DotLiquid;
using ExcelLibs;

namespace App
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Executing...");
            FileInfo templateFile = new FileInfo(args[0]);
            FileInfo outputFile = new FileInfo(args[1]);
            Console.WriteLine($"template file: {templateFile.FullName}");
            Console.WriteLine($"output file: {outputFile.FullName}");
            ExcelService excelService = new ExcelService();
            excelService.TEST(templateFile,outputFile,QTSC_Hash());
            Console.WriteLine("Complete");
//            Console.WriteLine("Press anykey to continue !");
//            Console.ReadKey();
        }

//        public static void Main(string[] args)
//        {
//            Template template = Template.Parse(File.ReadAllText("template.txt"));
//            var data = GetDataHash();
//            String rendered = template.Render(data);
//            Console.WriteLine("Press anykey to continue !");
//            Console.ReadKey();
//        }
//        

        private static Hash QTSC_Hash()
        {
            List<Item> itemsList1 = new List<Item>
            {
                new Item(1,"item a",new decimal(69.2),"tháng"),
                new Item(2,"item b",new decimal(88.3),"tháng"),
                new Item(3,"item c",new decimal(27.6),"tháng"),
                new Item(4,"item d",new decimal(63.20),"tháng"),
                new Item(5,"item abc",new decimal(115.7),"tháng")
            };
            var itemshash = itemsList1.Select(ObjectToHash);
            object obj = new
            {
                contactName = "contact",
                customerName = "customer",
                categoryGroup = "Thuê mặt bằng",
                items = itemshash,
                tax = 9.3,
                discount = 63,
                staffName = "TLP",
                phone = 97996541,
                email = "tlphong@gmail.com"
            };

            return ObjectToHash(obj);
        }
        

        private static Hash GetDataHash()
        {
            List<Item> itemsList1 = new List<Item>
            {
                new Item("item a"),
                new Item("item b"),
                new Item("item c"),
                new Item("item abc")
            };

            List<Item> itemsList2 = new List<Item>
            {
                new Item("item 1"),
                new Item("item 2"),
                new Item("item 3")
            };

            List<Category> cates =
                new List<Category> {new Category("cate a1", itemsList2), new Category("cate a2", itemsList1)};

            var itemshash = itemsList1.Select(ObjectToHash);
            var cateshash = cates.Select(cate => ObjectToHash(new
            {
                cate.name,
                items = cate.items.Select(ObjectToHash)
            }));

            object obj = new
            {
                id = 1,
                product = "abc",
                quantity = 2,
                price = 30.2,
                value = "some value",
                customerName = "cus name",
                customerEmail = "some email",
                items = itemshash,
                categories = cateshash
            };
            return ObjectToHash(obj);
        }

        private static Hash ObjectToHash(Object obj)
        {
            Hash resultHash = new Hash();
            Type type = obj.GetType();
            foreach (PropertyInfo property in type.GetProperties())
            {
                if (property.CanRead)
                {
                    resultHash[property.Name] = property.GetValue(obj);
                }
            }

            return resultHash;
        }
    }
}