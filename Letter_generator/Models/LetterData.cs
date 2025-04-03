namespace Letter_generator.Models
{
    public class LetterData
    {
        public string? CurrentLettNumb { get; set; }
        public DateOnly? CurrentDate { get; set; }
        public string? IncomingLettNumb { get; set; }
        public DateOnly? IncomingDate { get; set; }
        public string? Sex { get; set; }
        public string? Post { get; set; }
        public string? Organization { get; set; }
        public string? Recipient_FIO { get; set; }
        public string? Theme { get; set; }
        public string? Text { get; set; }
        public string? Performer_FIO { get; set; }
        public string? Performer_phone { get; set; }
        public string? Performer_email { get; set; }
        public List<Attachment> Attachments { get; set; }
    }   
}
