using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Word_Тепляков.Models;
using Microsoft.Office.Interop.Word;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using static PdfSharp.Capabilities.Features;
using Word_Тепляков.Elements;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Diagnostics;

namespace Word_Тепляков.Context
{
    public class OwnerContext : Models.Owner
    {
        public OwnerContext(string FirstName, string LastName, string SurName, int NumberRoom, BitmapImage Img, bool IsOwner) : base(FirstName, LastName, SurName, NumberRoom, Img, IsOwner) { }
    

        public static List<OwnerContext> AllOwners()
        {
            List<OwnerContext> allOwners = new List<OwnerContext>();
            allOwners.Add(new OwnerContext("Елена", "Иванова", "Петровна", 1, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Алексей", "Смирнов", "Владимирович", 2, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Анна", "Кузнецова", "Сергеевна", 3, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), false));
            allOwners.Add(new OwnerContext("Дмитрий", "Павлов", "Александрович", 3, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Ольга", "Михайловна", "Ивановна", 4, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Артем", "Козлов", "Олегович", 5, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Наталья", "Соколова", "Викторовна", 6, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Игорь", "Лебедев", "Андреевич", 6, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), false));
            allOwners.Add(new OwnerContext("Екатерина", "Федорова", "Дмитриевна", 7, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), false));
            allOwners.Add(new OwnerContext("Андрей", "Александров", "Игоревич", 7, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Оксана", "Степановна", "Николаевна", 8, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Сергей", "Никитин", "Васильевич", 9, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Мария", "Ковалева", "Александровна", 10, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Павел", "Фролов", "Михайлович", 11, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Елена", "Белова", "Александровна", 12, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Илья", "Поляков", "Данилович", 13, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Анастасия", "Гаврилова", "Валерьевна", 14, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Денис", "Орлов", "Владимирович", 15, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Алина", "Киселева", "Сергеевна", 16, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), false));
            allOwners.Add(new OwnerContext("Артем", "Ткаченко", "Викторович", 16, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), false));
            allOwners.Add(new OwnerContext("Валерия", "Романова", "Павловна", 16, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Александр", "Максимов", "Юрьевич", 17, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Евгения", "Сидорова", "Игоревна", 17, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), false));
            allOwners.Add(new OwnerContext("Никита", "Антонов", "Алексеевич", 18, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            allOwners.Add(new OwnerContext("Юлия", "Дмитриева", "Владимировна", 19, new BitmapImage(new Uri("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png")), true));
            return allOwners;
        }

        public static void Report(string fileName)
        {
            Application app = new Application();
            Document doc = app.Documents.Add();
            Paragraph paraHeader = doc.Paragraphs.Add();
            paraHeader.Range.Font.Size = 16;
            paraHeader.Range.Text = "Список жильцов дома";
            paraHeader.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            paraHeader.Range.Font.Bold = 1;
            paraHeader.Range.InsertParagraphAfter();
            Paragraph paraAddress = doc.Paragraphs.Add();
            paraAddress.Range.Font.Size = 14;
            paraAddress.Range.Text = "по адресу: г. Пермь, ул. Луначарского, д. 24";
            paraHeader.Range.ParagraphFormat.SpaceAfter = 20;
            paraHeader.Range.Font.Bold = 0;
            paraAddress.Range.InsertParagraphAfter();
            Paragraph paraCount = doc.Paragraphs.Add();
            paraCount.Range.Font.Size = 14;
            paraCount.Range.Text = $"Всего жильцов: {AllOwners().Count}";
            paraCount.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            paraCount.Range.InsertParagraphAfter();
            Paragraph paraTable = doc.Paragraphs.Add();
            Table paymentsTable = doc.Tables.Add(paraTable.Range, AllOwners().Count + 1, 6);
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            paymentsTable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Cell("№", paymentsTable.Cell(1, 1).Range);
            Cell("Фамилия", paymentsTable.Cell(1, 2).Range);
            Cell("Имя", paymentsTable.Cell(1, 3).Range);
            Cell("Отчество", paymentsTable.Cell(1, 4).Range);
            Cell("Квартира", paymentsTable.Cell(1, 5).Range);
            Cell("Изображение", paymentsTable.Cell(1, 6).Range);
            int temp1 = -1;
            int temp2 = -1;
            for (int i = 0; i < AllOwners().Count; i++)
            {
                OwnerContext owner = AllOwners()[i];
                Cell((i + 1).ToString(), paymentsTable.Cell(1 + 1 + i, 1).Range);
                Cell(owner.LastName, paymentsTable.Cell(1 + 1 + i, 2).Range, WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.FirstName, paymentsTable.Cell(1 + 1 + i, 3).Range, WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.SurName, paymentsTable.Cell(1 + 1 + i, 4).Range, WdParagraphAlignment.wdAlignParagraphLeft);
                if (owner.NumberRoom != temp1)
                {
                    Cell(owner.NumberRoom.ToString(), paymentsTable.Cell(2 + i, 5).Range);
                    temp2 = i + 2;
                }
                else 
                { 
                    paymentsTable.Cell(temp2, 5).Merge(paymentsTable.Cell(2 + i, 5)); 
                }
                temp1 = owner.NumberRoom;
                Cell(owner.Img, paymentsTable.Cell(1 + 1 + i, 6).Range, WdParagraphAlignment.wdAlignParagraphCenter);
            }
            doc.SaveAs2(fileName);
            doc.Close();
            app.Quit();
        }

        public static void Cell(string Text, Range Cell, WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphCenter)
        {
            Cell.Text = Text;
            Cell.ParagraphFormat.Alignment = alignment;
        }

        public static void Cell(BitmapImage Image, Range Cell, WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphCenter)
        {
            Cell.InlineShapes.AddPicture(Image.ToString());
            Cell.ParagraphFormat.Alignment = alignment;
        }

        public static void ReportPDF(string FileName)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Отчет по жильцам дома";
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            int MarginTop = 20;
            int MarginLeft = 50;
            XFont fontHeader = new XFont("Arial", 16, XFontStyleEx.Bold);
            XFont font = new XFont("Arial", 12);
            gfx.DrawString("Список жильцов дома", fontHeader, XBrushes.Black, new XRect(0, MarginTop, page.Width, 15), XStringFormats.Center);
            gfx.DrawString("по адресу: г. Пермь, ул. Луначарского, д. 24", font, XBrushes.Black, new XRect(0, MarginTop + 30, page.Width, 10), XStringFormats.Center);
            gfx.DrawString("Всего жильцов: " + AllOwners().Count, font, XBrushes.Black, new XRect(MarginLeft, MarginTop + 70, page.Width, 10), XStringFormats.CenterLeft);
            int Width = (Convert.ToInt32(page.Width.Value) - MarginLeft * 2 - 30) / 5;
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width + 10, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 2, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 3, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 4, MarginTop + 100, Width, 20);
            gfx.DrawString("№ квартиры", font, XBrushes.Black, new XRect(MarginLeft, MarginTop + 100, Width, 20), XStringFormats.Center);
            gfx.DrawString("Фамилия", font, XBrushes.Black, new XRect(MarginLeft + Width + 10, MarginTop + 100, Width, 20), XStringFormats.Center);
            gfx.DrawString("Имя", font, XBrushes.Black, new XRect(MarginLeft + (Width + 10) * 2, MarginTop + 100, Width, 20), XStringFormats.Center);
            gfx.DrawString("Отчество", font, XBrushes.Black, new XRect(MarginLeft + (Width + 10) * 3, MarginTop + 100, Width, 20), XStringFormats.Center);
            gfx.DrawString("Изображение", font, XBrushes.Black, new XRect(MarginLeft + (Width + 10) * 4, MarginTop + 100, Width, 20), XStringFormats.Center);
            PdfPage page2 = document.AddPage();
            XGraphics gfx2 = XGraphics.FromPdfPage(page2);
            gfx2.DrawString("Список собственников", fontHeader, XBrushes.Black, new XRect(0, MarginTop, page2.Width, 15), XStringFormats.Center);
            gfx2.DrawString("по адресу: г. Пермь, ул. Луначарского, д. 24", font, XBrushes.Black, new XRect(0, MarginTop + 30, page.Width, 10), XStringFormats.Center);
            int Width2 = (Convert.ToInt32(page2.Width.Value) - MarginLeft * 2 - 30) / 4;
            gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 60, Width2, 20);
            gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width2 + 10, MarginTop + 60, Width2, 20);
            gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width2 + 10) * 2, MarginTop + 60, Width2, 20);
            gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width2 + 10) * 3, MarginTop + 60, Width2, 20);
            gfx2.DrawString("№ квартиры", font, XBrushes.Black, new XRect(MarginLeft, MarginTop + 60, Width2, 20), XStringFormats.Center);
            gfx2.DrawString("Фамилия", font, XBrushes.Black, new XRect(MarginLeft + Width2 + 10, MarginTop + 60, Width2, 20), XStringFormats.Center);
            gfx2.DrawString("Имя", font, XBrushes.Black, new XRect(MarginLeft + (Width2 + 10) * 2, MarginTop + 60, Width2, 20), XStringFormats.Center);
            gfx2.DrawString("Отчество", font, XBrushes.Black, new XRect(MarginLeft + (Width2 + 10) * 3, MarginTop + 60, Width2, 20), XStringFormats.Center);
            int temp1 = -1;
            int temp2 = -1;
            int tempOwners = 0;
            for (int i = 0; i < AllOwners().Count; i++)
            {
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width + 10, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 2, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 3, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 4, MarginTop + 100 + 25 * (i + 1), Width, 20);
                if (AllOwners()[i].NumberRoom != temp1)
                {
                    gfx.DrawString(AllOwners()[i].NumberRoom.ToString(), font, XBrushes.Black, new XRect(MarginLeft, MarginTop + 100 + 25 * (i + 1), Width, 20), XStringFormats.Center);
                    temp2 = i + 2;
                }
                else gfx.DrawString("", font, XBrushes.Black, new XRect(MarginLeft, MarginTop + 100 + 25 * (i + 1), Width, 20), XStringFormats.Center);
                temp1 = AllOwners()[i].NumberRoom;
                gfx.DrawString(AllOwners()[i].LastName, font, XBrushes.Black, new XRect(MarginLeft + Width + 10, MarginTop + 100 + 25 * (i + 1), Width, 20), XStringFormats.Center);
                gfx.DrawString(AllOwners()[i].FirstName, font, XBrushes.Black, new XRect(MarginLeft + (Width + 10) * 2, MarginTop + 100 + 25 * (i + 1), Width, 20), XStringFormats.Center);
                gfx.DrawString(AllOwners()[i].SurName, font, XBrushes.Black, new XRect(MarginLeft + (Width + 10) * 3, MarginTop + 100 + 25 * (i + 1), Width, 20), XStringFormats.Center);
                XImage image = XImage.FromFile("C:\\Users\\kiril\\Desktop\\MDK_01_01_PR50\\Images\\owner.png");
                gfx.DrawImage(image, new XRect(MarginLeft + (Width + 10) * 4 + 35, MarginTop + 100 + 25 * (i + 1), 20, 20));
                if(AllOwners()[i].IsOwner == true)
                {
                    gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20);
                    gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width2 + 10, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20);
                    gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width2 + 10) * 2, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20);
                    gfx2.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width2 + 10) * 3, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20);
                    gfx2.DrawString(AllOwners()[i].NumberRoom.ToString(), font, XBrushes.Black, new XRect(MarginLeft, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20), XStringFormats.Center);
                    gfx2.DrawString(AllOwners()[i].LastName, font, XBrushes.Black, new XRect(MarginLeft + Width2 + 10, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20), XStringFormats.Center);
                    gfx2.DrawString(AllOwners()[i].FirstName, font, XBrushes.Black, new XRect(MarginLeft + (Width2 + 10) * 2, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20), XStringFormats.Center);
                    gfx2.DrawString(AllOwners()[i].SurName, font, XBrushes.Black, new XRect(MarginLeft + (Width2 + 10) * 3, MarginTop + 60 + 25 * (tempOwners + 1), Width2, 20), XStringFormats.Center);
                    tempOwners++;
                }
            }
            document.Save(FileName);
            Process.Start(FileName);
        }
    }
}
