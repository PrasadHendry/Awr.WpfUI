namespace Awr.Core.DTOs
{
    // DTO for collecting a single line item's data from the UI before submission
    public class AwrItemSubmissionDto
    {
        public string MaterialProduct { get; set; }
        public string BatchNo { get; set; }
        public string ArNo { get; set; }
        public string AwrNo { get; set; }
        public decimal QtyRequired { get; set; }
    }
}