using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Xml.Linq;
using AutoMapper;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace SharePointListToPDF
{
    internal class Program
    {
        static Program()
        {
            Mapper.CreateMap<XElement, ListItem>()
                  .ForMemberFromSPXElement(x => x.ID)
                  .ForMemberFromSPXElement(x => x.LinkTitle)
                  .ForMemberFromSPXElement(x => x.Status)
                  .ForMemberFromSPXElement(x => x.BoardReport, "ows_Board_x0020_Report");
        }

        private static void Main(string[] args)
        {
            var baseDirectoryPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            
            var list = GetList<ListItem>();
            var table = CreatePdfTable(list);
            WritePdf(baseDirectoryPath, table);
        }

        private static IEnumerable<TModel> GetList<TModel>()
        {
            return SPListHelper.GetListItems().Select(Mapper.Map<XElement, TModel>).ToArray();
        }

        private static PdfPTable CreatePdfTable(IEnumerable<ListItem> list)
        {
            var table = new PdfPTable(3) { TotalWidth = 500f };
            table.SetWidths(new[] { 225f, 50f, 225f });
            var cell = new PdfPCell(new Phrase("Board Report"))
            {
                Colspan = 3,
                HorizontalAlignment = 1
            };
            table.AddCell(cell);
            table.AddCell("Charge");
            table.AddCell("Status");
            table.AddCell("Board Report");
            foreach (var listItem in list)
            {
                table.AddCell(listItem.LinkTitle);
                PdfPCell statusCell = GetStatusCell(listItem.Status);
                table.AddCell(statusCell);
                table.AddCell(listItem.BoardReport);
            }
            return table;
        }

        private static void WritePdf(string baseDirPath, PdfPTable table)
        {
            var doc = new Document();
            var path = GetPdfPath(baseDirPath);
            PdfWriter.GetInstance(doc, new FileStream(path, FileMode.Create));
            doc.Open();
            doc.AddTitle("Board Report");
            doc.AddSubject("Board Report");
            //doc.Add(new Paragraph("Board Report"));
            doc.Add(table);
            doc.Close();
        }

        private static PdfPCell GetStatusCell(string status)
        {
            var statusCell = new PdfPCell();
            var color = BaseColor.BLUE; // "Completed"
            if (status == "Halted or significant problems encountered")
            {
                color = BaseColor.RED;
            }
            else if (status == "Slowed or off-track")
            {
                color = BaseColor.YELLOW;
            }
            else if (status == "Underway, on target")
            {
                color = BaseColor.GREEN;
            }
            statusCell.BackgroundColor = color;
            return statusCell;
        }

        private static string GetPdfPath(string appDomainBaseDir)
        {
            var path = Path.Combine(appDomainBaseDir, "reports");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return Path.Combine(path, "report" + new Random().Next() + ".pdf");
        }

// ReSharper disable ClassNeverInstantiated.Local
        class ListItem
// ReSharper restore ClassNeverInstantiated.Local
        {
            public string ID { get; set; }

            public string LinkTitle { get; set; }

            public string Status { get; set; }

            public string BoardReport { get; set; }
        }
    }

    public static class AutoMapperExtensions
    {
// ReSharper disable InconsistentNaming
        public static IMappingExpression<XElement, TDestination> ForMemberFromSPXElement<TDestination>(
// ReSharper restore InconsistentNaming
            this IMappingExpression<XElement, TDestination> mappingExpression,
            Expression<Func<TDestination, object>> destinationMember, string attributeName = null)
        {
            var memberExpr = destinationMember.Body as MemberExpression;
            attributeName = attributeName ?? "ows_" + memberExpr.Member.Name;
            return mappingExpression
                .ForMember(destinationMember,
                           opt => opt.ResolveUsing(elem => elem.Attribute(XName.Get(attributeName)).Value));
        }
    }
}
