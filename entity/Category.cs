using System;
using System.Collections.Generic;

namespace App
{
    public class Category
    {
        public String name { get; set; }
        public List<Item> items { get; set; }

        public Category(string name, List<Item> items)
        {
            this.name = name;
            this.items = items;
        }

        public Category(string name)
        {
            this.name = name;
        }
    }
}