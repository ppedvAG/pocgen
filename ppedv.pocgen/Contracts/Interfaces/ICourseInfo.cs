using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Interfaces
{
    public interface ICourseInfo
    {
        string CourseName { get; set; }
        string CourseCurrentModuleName { get; set; }
    }
}
