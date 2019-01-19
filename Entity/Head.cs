using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Entity
{
    public class Head : BaseHead
    {
        public string School { get; set; }
        public string Class { get; set; }
        public string Semester { get; set; }
        public string TeacherInCharge { get; set; }
        public DateTime? OpenDate { get; set; }
        public decimal? XueFei { get; set; }
    }
}
