using MessagePack;
using System.ComponentModel.DataAnnotations;
using KeyAttribute = System.ComponentModel.DataAnnotations.KeyAttribute;

namespace InchesExcel.Models
{
    public class ClientExcel
    {        
        [Key]
        public int Id { get; set; }

        [Display(Name = "policy number")]
        public string? Policynumber { get; set; }
        
        [Display(Name = "UW Name")]
        public string? UWName { get; set; }
        
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString ="{0:yyyy-MM-dd}")]

        public DateTime? Date { get; set; }

        public string? Lot { get; set; }

        [Display(Name = "Received timing")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:HH:mm}")]

        public DateTime? Receivedtiming { get; set; }

        [Display(Name = "Done Cases")]
        
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:HH:mm}")]

        public DateTime? DoneCases { get; set; }

        public int TAT { get; set; }

        [Display(Name = "With in TAT")]
        public string? WithinTAT { get; set; }

        [Display(Name = "Cases Status")]
        public string? CasesStatus { get; set; }


    }
}
