using pocgen.Contracts.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Models
{
    public class CourseInfo : ICourseInfo
    {
        public string CourseName { get; set; }
        public string CourseCurrentModuleName { get; set; }
    }
}
