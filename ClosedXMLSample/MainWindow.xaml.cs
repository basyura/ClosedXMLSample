using System;
using System.IO;
using System.Windows;
using ClosedXML.Excel;

namespace ClosedXMLSample
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Contacts");

            ws.Cell("B2").Value = "Contacts";

            ws.Cell("B3").Value = "FName";
            ws.Cell("B4").Value = "Jhon";
            ws.Cell("B5").Value = "Hank";
            ws.Cell("B6").SetValue("Dangy");


            // Last Names
            ws.Cell("C3").Value = "LName";
            ws.Cell("C4").Value = "Galt";
            ws.Cell("C5").Value = "Rearden";
            ws.Cell("C6").SetValue("Taggart"); // Another way to set the value


            // Boolean
            ws.Cell("D3").Value = "Outcast";
            ws.Cell("D4").Value = true;
            ws.Cell("D5").Value = false;
            ws.Cell("D6").SetValue(false); // Another way to set the value

            // DateTime
            ws.Cell("E3").Value = "DOB";
            ws.Cell("E4").Value = new DateTime(1919, 1, 21);
            ws.Cell("E5").Value = new DateTime(1907, 3, 4);
            ws.Cell("E6").SetValue(new DateTime(1921, 12, 15)); // Another way to set the value

            // Numeric
            ws.Cell("F3").Value = "Income";
            ws.Cell("F4").Value = 2000;
            ws.Cell("F5").Value = 40000;
            ws.Cell("F6").SetValue(10000); // Another way to set the value

            var rngTable = ws.Range("B2:F6");

            var rngDates = rngTable.Range("D3:D5");
            var rngNumbers = rngTable.Range("E3:E5");

            rngDates.Style.NumberFormat.NumberFormatId = 15;
            rngNumbers.Style.NumberFormat.Format = "$ $,$$0";

            rngTable.FirstCell().Style
                .Font.SetBold()
                .Fill.SetBackgroundColor(XLColor.CornflowerBlue)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            rngTable.FirstRow().Merge();

            var rngData = ws.Range("B3:F6");
            var excelTable = rngData.CreateTable();

            excelTable.ShowTotalsRow = true;
            excelTable.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Average;
            excelTable.Field("DOB").TotalsRowLabel = "Average:";

            ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            ws.Columns().AdjustToContents();

            string path =  Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),"ClosedXmlSample.xlsx");

            wb.SaveAs(path);
        }
    }
}
