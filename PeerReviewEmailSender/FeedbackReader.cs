using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PeerReviewEmailSender
{
    public class FeedbackReader
    {
        private readonly EmailContainer targetContainer;

        public FeedbackReader(EmailContainer targetContainer)
        {
            this.targetContainer = targetContainer;
        }

        public void ReadFeedbacks(string xlsFilename)
        {
            using (FileStream fileStream = new FileStream(xlsFilename, FileMode.Open))
            {
                ExcelPackage excel = new ExcelPackage(fileStream);
                var workSheet = excel.Workbook.Worksheets[1];

                IEnumerable<ExcelRowDto> newcollection =
                    workSheet.ConvertSheetToObjects<ExcelRowDto>();
                newcollection.ToList().ForEach(
                    f => targetContainer.Add(f.Email, f.ToFeedbackString));
            }
        }

        class ExcelRowDto
        {
            [Column(1)]
            public string Timestamp { get; set; }

            [Column(2)]
            public string Email { get; set; }

            [Column(3)]
            public string PresentationScore { get; set; }

            [Column(4)]
            public string WorkScore { get; set; }

            [Column(5)]
            public string Feedback { get; set; }

            public string ToFeedbackString =>
                $"Előadás: {PresentationScore}{Environment.NewLine}<BR/>Munka: {WorkScore}{Environment.NewLine}{Environment.NewLine}<BR/>Visszajelzés:{Environment.NewLine}<BR/>{Feedback}";
        }
    }
}
