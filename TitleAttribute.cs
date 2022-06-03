﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Advanced.Common
{
    [AttributeUsage(AttributeTargets.Property)]
    public class TitleAttribute:Attribute
    {
        public string Title;
    }
}
