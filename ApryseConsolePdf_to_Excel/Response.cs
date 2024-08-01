using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApryseConsolePdf_to_Excel
{
    public class Response
    {
        public string? FileName { get; set; }
        public string? InputUrl { get; set; }
        public string? OutputUrl { get; set; }
        public string? ErrorMessage { get; set; }
        public bool IsSuccess { get; set; }
    }
}
