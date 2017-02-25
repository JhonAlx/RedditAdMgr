﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RedditAdMgr.Model
{
    class Campaign
    {
        public Advertisement Advertisement { get; set; }
        public string Target { get; set; }
        public string TargetDetail { get; set; }
        public string Location { get; set; }
        public string Location2 { get; set; }
        public string Platform { get; set; }
        public decimal Budget { get; set; }
        public bool BudgetOptionDeliverFast { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public bool OptionExtend { get; set; }
        public decimal PricingCPM { get; set; }
    }
}
