using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace Letter_generator.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LetterDataController : ControllerBase
    {

        [HttpPost("generate")]
        public IActionResult GenerateLetter([FromBody] Models.LetterData data)
        {
            if (data == null) return BadRequest("Нет данных");

            string templatePath = "template1.docx";
            if (!System.IO.File.Exists(templatePath))
            {
                return NotFound("Шаблон не найден");
            }
            string outputPath = Path.Combine(Path.GetTempPath(), $"letter_{Guid.NewGuid()}.docx");

            System.IO.File.Copy(templatePath, outputPath, true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                foreach (var text in body.Descendants<Text>())
                {
                    text.Text = text.Text
                        .Replace("{CurrentLettNumb}", data.CurrentLettNumb)
                        .Replace("{CurrentDate}", data.CurrentDate)
                        .Replace("{IncomingLettNumb}", data.IncomingLettNumb)
                        .Replace("{IncommingDate}", data.IncommingDate)
                        .Replace("{Post}", data.Post)
                        .Replace("{Organization}", data.Organization)
                        .Replace("{Sender_FIO}", data.Sender_FIO)
                        .Replace("{Theme}", data.Theme)
                        .Replace("{Text}", data.Text)
                        .Replace("{Performer_FI}", data.Performer_FIO)
                        .Replace("{Performer_phone}", data.Performer_phone)
                        .Replace("{Performer_email}", data.Performer_email);
                }
                doc.MainDocumentPart.Document.Save();
            }

            byte[] fileBytes = System.IO.File.ReadAllBytes(outputPath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "letter.docx");
        }
    }

}


