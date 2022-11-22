using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
namespace ExcelApp.Models
{
    public class UploadModel
    {
        [DisplayName("Upload File")]
        public IFormFile? ExcelFile { get; set; }
        
    }
}
