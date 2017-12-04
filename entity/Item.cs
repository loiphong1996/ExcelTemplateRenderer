using System;

namespace App
{
    public class Item
    {
        public int count { get; set; }
        public String name { get; set; }
        public decimal price { get; set; }
        public String unit { get; set; }

        public Item(string name)
        {
            this.name = name;
        }

        public Item(int count, string name, decimal price, string unit)
        {
            this.count = count;
            this.name = name;
            this.price = price;
            this.unit = unit;
        }
    }
}