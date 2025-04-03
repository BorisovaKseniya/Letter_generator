using Letter_generator.Models;
using Letter_generator.wwwroot.Services;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics.Metrics;


namespace Letter_generator.Controllers
{
    [Route("/generate")]
    [ApiController]
    public class Controller : ControllerBase
    {
        private readonly Service _service;

        public Controller(Service service)
        {
            _service = service;
        }

        [HttpPost]
        public IActionResult ReceiveData([FromForm] LetterData input)
        {
            _service.SaveData(input);

            var letter = _service.GetData();

            if (letter == null)
                return BadRequest("Данные для письма отсутствуют");

            var fileBytes = _service.CreateDocument();

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"Super Letter.docx");
        }
    }
}
