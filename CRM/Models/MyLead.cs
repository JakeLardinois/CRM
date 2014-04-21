using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Xrm;

namespace CRM.Models
{
    /*I created this lead class that inherits from the Lead class so that I could populate the survey data into a note in my Leads collection*/
    class MyLead : Lead
    {
        public Annotation Note { get; set; }
    }
}
