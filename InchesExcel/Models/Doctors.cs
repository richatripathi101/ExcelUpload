using System.ComponentModel.DataAnnotations;

namespace InchesExcel.Models
{
    public class Doctors
    {
        [Key]
        public int Id { get; set; }

        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime? Date { get; set; }
        public String? PolicyNumber { get; set; }
        public String? DecisionMatchWithInitialUW { get; set; }
        public String? Decision { get; set; }
        public String? Remarks { get; set; }

        public String? CPDetailAndFindings { get; set; }
        public String? InchesUWname { get; set; }

        public String? FinalDecision { get; set; }

        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:HH:mm}")]
        public DateTime? LOTsenttime { get; set; }

        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:HH:mm}")]
        public DateTime? Receivedtime { get; set; }

        public String? TATstatus { get; set; }

        public String? Team { get; set; }

        public String? SystemStatus { get; set; }

        public String? CaseType { get; set; }
       

    }
}
