using ClassLibraryNetCore;
using Docentric.Documents.ObjectModel;
using Docentric.Documents.Reporting;
using Docentric.Drawing;
using System;
using System.IO;

namespace DocentricConsoleApplicationNetCore
{
    public class ImageGeneration : ReportBase
    {
        public ImageGeneration()
        {
            Name = "Image";
        }

        public override string GenerateReport()
        {
            ImageData imageData = new()
            {
                Data = File.ReadAllBytes(Path.Combine(ApplicationPath, "Resources\\" + Name + "\\HtmlFragment.html"))
            };

            // Create a temporary file for the generated report document.
            string reportDocumentFileName = Path.GetTempPath() + "GeneratedReport_" + Guid.NewGuid().ToString() + OutputDocumentFileExtension;

            using (Stream reportDocumentStream = File.Create(reportDocumentFileName))
            {
                // Open the report template file.
                using Stream reportTemplateStream = GetReportTemplate();

                // Generate the report document using 'DocumentGenerator'.
                DocumentGenerator dg = new(imageData);
                FontManager.SetFontsSources(new FolderFontSource("C:\\Windows\\Fonts"));

                DocumentGenerationResult result = dg.GenerateDocument(reportTemplateStream, reportDocumentStream, SaveOptions.Word);

                if (result.HasErrors)
                {
                    foreach (Error error in result.Errors)
                    {
                        Console.Out.WriteLine(error.Message);
                    }
                }
            }

            return reportDocumentFileName;
        }
    }
}