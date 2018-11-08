using ppedv.pocgen.Domain.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppedv.pocgen.Domain.Models
{
    public class CourseInfo : ICourseInfo
    {
        public string CourseName { get; set; }
        public string CourseCurrentModuleName { get; set; }
    }
}
