using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Letter_generator.Models;
using System.Diagnostics.Metrics;

namespace Letter_generator.wwwroot.Services
{
    public class Service
    {
        private LetterData? _letter;
        public string recipient;

        public void SaveData(LetterData letter)
        {
            _letter = letter;
        }

        public LetterData? GetData()
        {
            return _letter;
        }

        public byte[] CreateDocument()
        {
            if (_letter == null)
            {
                throw new InvalidOperationException("Нет данных для создания письма");
            }

            using (var stream = new MemoryStream())
            {
                string template = Path.Combine(Directory.GetCurrentDirectory(), "Template1.docx");
                if (!File.Exists(template))
                    throw new InvalidOperationException("Нет данных для создания письма");

                using (FileStream fileStream = new FileStream(template, FileMode.Open, FileAccess.Read))
                {
                    fileStream.CopyTo(stream);
                }

                using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true))
                {
                    var body = document.MainDocumentPart.Document.Body;

                    string[] recipient = _letter.Recipient_FIO.Split(' ');
                    string[] sender = _letter.Performer_FIO.Split(' ');
                    this.recipient = recipient.Length >= 3 ? recipient[0] + ' ' + recipient[1][0] + ". " + recipient[2][0] + ". " : _letter.Recipient_FIO;

                    var placeholders = new Dictionary<string, string>
                    {
                        { "{CurNumb}", _letter.CurrentLettNumb },
                        { "{CurDate}", _letter.CurrentDate.ToString() },
                        { "{Post}", _letter.Post },
                        { "{IncDate}", _letter.IncomingDate != null ? "на " + _letter.IncomingDate : "" },
                        { "{IncNumb}", _letter.IncomingLettNumb != null ? "№" + _letter.IncomingLettNumb : "" },
                        { "{Organization}", _letter.Organization },
                        { "{RecFIO}", this.recipient},
                        { "{Theme}", _letter.Theme },
                        { "{Text}", _letter.Text },
                        { "{Recipient}", (_letter.Sex == "male" ? "Уважаемый " : "Уважаемая ") + recipient[0] + ' ' + (recipient.Length >= 2 ? recipient[1] : "") },
                        { "{Attachment}", ""}
                    };


                    if (_letter.Attachments.Count == 0)
                    {
                        placeholders["{Attachment}"] = "Приложения:";
                        int i = 1;
                        foreach (var attachment in _letter.Attachments)
                        {
                            body.AppendChild(CreateParagraph($"{i++}. {attachment.Theme}"));
                        }
                        i = 1;
                        foreach (var attachment in _letter.Attachments)
                        {

                            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page }))); // Разрыв страницы
                            body.AppendChild(CreateRightAlignedParagraph($"Приложение {i}"));
                            body.AppendChild(CreateCenterAlignedParagraph(attachment.Theme));
                            body.AppendChild(CreateParagraph(attachment.Text));
                            i++;
                        }
                    }

                    foreach (var text in body.Descendants<Text>())
                    {
                        Console.WriteLine($"Текст в документе: '{text.Text}'");
                    }

                    var textElements = body.Descendants<Text>().Where(t => placeholders.Keys.Any(p => t.Text.Contains(p)));

                        foreach (var text in textElements)
                    {
                        foreach (var placeholder in placeholders)
                        {
                            if (text.Text.Contains(placeholder.Key))
                            {
                                text.Text = text.Text.Replace(placeholder.Key, placeholder.Value);
                            }
                        }
                    }

                    foreach (var footer in document.MainDocumentPart.FooterParts)
                    {
                        foreach (var text in footer.RootElement.Descendants<Text>())
                        {
                            text.Text = text.Text.Replace("{Performer_FIO}", (sender.Length >= 3 ? sender[0] + ' ' + sender[1][0] + ". " + sender[2][0] + ". " : _letter.Performer_FIO))
                                .Replace("{Performer_phone}", _letter.Performer_phone)
                                .Replace("{Performer_email}", _letter.Performer_email);
                        }
                    }

                    document.MainDocumentPart.Document.Save();

                }

                return stream.ToArray();
            }

        }
        private Paragraph CreateParagraph(string text)
        {
            return new Paragraph(new Run(new Text(text)));
        }

        private Paragraph CreateRightAlignedParagraph(string text)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Right } // Выравнивание по правому краю
                ),
                new Run(new Text(text))
            );
        }

        private Paragraph CreateCenterAlignedParagraph(string text)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center } // Выравнивание по правому краю
                ),
                new Run(
                    new RunProperties(
                        new Bold()),
                                        new Text(text))
            );
        }
    }
}
