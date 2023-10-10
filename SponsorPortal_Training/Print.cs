using System;
using System.IO;
using System.Data;
using System.Text;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Reporting.WebForms;
using NLog;
using SponsorPortal_Training.AppCode;

namespace SponsorPortal_Training
{
    class Printer : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private int m_currentPageIndex;
        private IList<Stream> m_streams;
        private Stream CreateStream(string name,
          string fileNameExtension, Encoding encoding,
          string mimeType, bool willSeek)
        {
            Stream stream = new MemoryStream();
            m_streams.Add(stream);
            return stream;
        }
        private void Export(LocalReport report)
        {
            string deviceInfo =
              @"<DeviceInfo>
                <OutputFormat>EMF</OutputFormat>
                <PageWidth>8.5in</PageWidth>
                <PageHeight>11in</PageHeight>
                <MarginTop>0in</MarginTop>
                <MarginLeft>0in</MarginLeft>
                <MarginRight>0in</MarginRight>
                <MarginBottom>0in</MarginBottom>
            </DeviceInfo>";
            Warning[] warnings;
            m_streams = new List<Stream>();
            report.Render("Image", deviceInfo, CreateStream,
               out warnings);
            foreach (Stream stream in m_streams)
                stream.Position = 0;
        }
        private void PrintPage(object sender, PrintPageEventArgs ev)
        {
            Metafile pageImage = new
               Metafile(m_streams[m_currentPageIndex]);

            // Adjust rectangular area with printer margins.
            Rectangle adjustedRect = new Rectangle(
                ev.PageBounds.Left - (int)ev.PageSettings.HardMarginX,
                ev.PageBounds.Top - (int)ev.PageSettings.HardMarginY,
                ev.PageBounds.Width,
                ev.PageBounds.Height);

            // Draw a white background for the report
            ev.Graphics.FillRectangle(Brushes.White, adjustedRect);

            // Draw the report content
            ev.Graphics.DrawImage(pageImage, adjustedRect);

            // Prepare for the next page. Make sure we haven't hit the end.
            m_currentPageIndex++;
            ev.HasMorePages = (m_currentPageIndex < m_streams.Count);
        }
        private void Print()
        {
            if (m_streams == null || m_streams.Count == 0)
                throw new Exception("Error: no stream to print.");

            PrintDocument printDoc = new PrintDocument();
            
            printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
            printDoc.PrinterSettings.PrinterName = Settings.Instance.Printer;
            _logger.Info("Printer name: " + Settings.Instance.Printer);
            m_currentPageIndex = 0;

            if (printDoc.PrinterSettings.IsValid)
            {
                _logger.Info("PrinterSettings is valid");
                printDoc.Print();
            }
            else
            {
                _logger.Info("PrinterSettings is not valid");
                _logger.Info("Exiting app");
                Environment.Exit(1);

            }
        }
        public void Run(List<GuestAccount> listGuestAccounts)
        {
            _logger.Info("Run printing guestAccounts. Number accounts to print: " + listGuestAccounts.Count);
            LocalReport report = new LocalReport();
            report.ReportPath = @"Report\UserReport.rdlc";
            report.DataSources.Add(new ReportDataSource("DataSet1", listGuestAccounts));

            _logger.Info("ReportPath: " + report.ReportPath);

            Export(report);
            Print();
        }
        public void Dispose()
        {
            if (m_streams != null)
            {
                foreach (Stream stream in m_streams)
                    stream.Close();
                m_streams = null;
            }
        }
    }
}
