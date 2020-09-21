using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelGenerator.Models
{
    public class TestModel
    {
        public int TestId { get; set; }
        public string TestName { get; set; }
        public string TestDesc { get; set; }
        public DateTime TestDate { get; set; }
    }
    public class TestModelList
    {
        public List<TestModel> TestData { get; set; }
    }
}
