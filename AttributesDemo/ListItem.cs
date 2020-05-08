using System;
using System.Collections.Generic;
using System.Text;

namespace ItemSpecificsDemo
{
    public class ListItem
    {
        private string _Name;
        private string _Value;

        public ListItem()
        { }
        
        public ListItem(string name, string value)
        {
            _Name = name;
            _Value = value;
        }

        public string Name
        {
            get { return _Name; }
            set { _Name = value; }
        }

        public string Value
        {
            get { return _Value; }
            set { _Value = value; }
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
