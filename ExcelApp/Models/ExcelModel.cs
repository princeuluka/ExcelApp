using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace ExcelApp.Models
{
    public class ExcelModel
    {
        [Required]
        [Key]
        public int Identity { get; set; }
        [Required]
        [DisplayName("First Name")]
        public string FirstName { get; set; }=String.Empty;
        [Required]
        public string Surname { get; set; } = String.Empty;
        [Required]
        public string Age { get; set; } = String.Empty;
        [Required]
        public string Sex { get; set; } = String.Empty;
        [Required]
        public string Mobile { get; set; } = String.Empty;
        [Required]
        public string Active { get; set; } = String.Empty;  
    }
}
