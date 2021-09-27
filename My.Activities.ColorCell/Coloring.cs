using System;
using System.Activities;
using System.ComponentModel;
using System.IO;
using ClosedXML.Excel;
using System.Drawing;
using System.Text.RegularExpressions;

namespace My.Activities.ColorCell
{
    public class Coloring : CodeActivity
    {

        private IXLWorkbook book;
        private IXLWorksheet worksheet;

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> PathExcel { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> SheetName { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Cell { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public System.Drawing.Color Color { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string path = Regex.Replace(PathExcel.Get(context), @"[^\P{C}\n]+","");
            string sheetName = SheetName.Get(context);
            string cell = Cell.Get(context);
            SetTip(path, sheetName, cell, Color);
        }

        public void SetTip(string path, string sheetName, string cell, System.Drawing.Color color)
        {
            if (File.Exists(path))
            {
                book = new XLWorkbook(path);

                try
                {
                    worksheet = book.Worksheet(sheetName);
                }
                catch
                {
                    worksheet = book.AddWorksheet(sheetName);
                }
            }
            else
            {
                book = new XLWorkbook();
                worksheet = book.AddWorksheet(sheetName);
            }

            if (!cell.Contains(":"))
                TipCell(cell, color);
            else
                TipRange(cell, color);
            book.Save();
            //book.SaveAs(path);
        }

        private void TipCell(string target, System.Drawing.Color color)
        {
            worksheet.Cell(target).Style.Fill.BackgroundColor = XLColor.FromColor(color) ;
        }
        private void TipRange(string target, System.Drawing.Color color)
        {
            string[] range = target.Split(':');

            IXLRange rangeXL;
            if (string.IsNullOrWhiteSpace(range[1]))
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), GetAlfb(worksheet.RangeUsed().FirstRowUsed().CellCount() - 1) + (worksheet.RangeUsed().RowCount()));
            }
            else
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), range[1].ToUpper());
            }
            rangeXL.Style.Fill.BackgroundColor = XLColor.FromColor(color);
        }
        private string GetAlfb(int num)
        {
            return (065 + num) > 90 ? ((char)Math.Floor(64 + (64.0 + num) / 90)).ToString() + ((char)(num % 90)).ToString() : ((char)(065 + num)).ToString();
        }
    }
}