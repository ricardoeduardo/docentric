using Docentric.Documents.ObjectModel;
using System;
using System.IO;

namespace DocentricConsoleApplicationNetCore
{
    public abstract class ReportBase
    {
        public string? Name { get; protected set; }
        public string? Title { get; protected set; }
        public string? Description { get; protected set; }
        public string? HelpTopicPage { get; protected set; }
        public string? DirectoryPath { get; private set; }

        // Constructor
        public ReportBase()
        {
            // Default output format is docx.
            SaveOptions = SaveOptions.Word;
        }

        /// <summary>
        /// Generates the report for the current template.
        /// </summary>
        /// <returns>File full name of the genarated report document.</returns>
        public abstract string GenerateReport();


        /// <summary>
        /// Opens the report template file for this example.
        /// </summary>
        /// <returns>Returns the file stream of the report template file.</returns>
        protected Stream GetReportTemplate()
        {
            if (!File.Exists(TemplateFileName))
            {
                throw new Exception(string.Format("Report template '{0}' is not found", TemplateFileName));
            }

            return File.Open(TemplateFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }


        // ApplicationPath
        protected string ApplicationPath
        {
            get
            {
                string? url = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);

                string? mainpath = url?.Replace("\\bin\\Debug", "")
                    .Replace("file:\\", "")
                    .Replace("netcoreapp3.1", "")
                    .Replace("net5.0", "")
                    .Replace("net6.0", "")
                    .Replace("net48", "");

                return mainpath ?? "";
            }
        }


        // TemplateFileName
        public string TemplateFileName => Path.Combine(ApplicationPath, "Resources\\" + Name + "\\Template.docx");

        // SaveOptions
        public SaveOptions SaveOptions {get; set; }

        protected string OutputDocumentFileExtension => SaveOptions switch
        {
            WordSaveOptions => ".docx",
            PdfSaveOptions => ".pdf",
            XpsSaveOptions => ".xps",
            _ => throw new Exception("Unsupported 'SaveOptions' object.")
        };
    }
}