using System.Runtime.InteropServices;

namespace MarketingCourseTool.Model
{
    public class StudentData
    {
        public int Index { get; set; }
        public string StudentId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public int Group { get; set; }
        public string DocUrl1 { get; set; } = " ";

        public string DocUrl2 { get; set; } = " ";
        public int[] Indexes { get; set; } = {0, 0};
        public string Message{ get; set; } = " ";
    }
}
