using System;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using OfficeOpenXml;
using System.Text;
using System.Drawing;


namespace MailToExcel
{
    class Program
    {

        public static class common { public static int currentNum { get; set; } }

        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 或者其他適合你的授權內容
            string msgFilePath = @"D:\no1.msg";
            string outputFilePath = @"D:\output.xlsx";

            var msg = new MsgReader.Outlook.Storage.Message(msgFilePath);

            // 檢查郵件的正文是否是 HTML 格式
            if (msg.BodyHtml != null)
            {
                var htmlContent = msg.BodyHtml;

                // 使用 HtmlAgilityPack 解析 HTML 內容
                var doc = new HtmlDocument();
                doc.LoadHtml(htmlContent);

                // 尋找所有的表格
                var tables = doc.DocumentNode.Descendants("table");

                // 建立 Excel 文件
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Table Data");

                    // 指定欄位名稱
                    string[] columns = { "PO Number", "Contact Person", "Currency", "Item", "Material", "Order qty.", "Price Base", "Price", "Net value", "Delivery date" };


                    // 在 Excel 中寫入欄位名稱並設定背景色
                    for (int i = 0; i < columns.Length; i++)
                    {
                        var cell = worksheet.Cells[1, i + 1];
                        cell.Value = columns[i];
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 85, 140));
                        cell.Style.Font.Color.SetColor(Color.White);
                    }

                    // 從第二行開始寫入表格數據
                    int rowNumber = 2;

                    // 遍歷每個表格
                    foreach (var table in tables)
                    {
                        // 獲取表格的所有行
                        var rows = table.SelectNodes("tr");
                        if (rows != null)
                        {
                            // 遍歷表格的每一行
                            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                            {
                                var row = rows[rowIndex];

                                // 獲取該行的所有單元格
                                var cells = row.SelectNodes("th|td");
                                if (cells != null)
                                {
                                    // 遍歷每個單元格，尋找指定欄位
                                    for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                                    {
                                        var cell = cells[cellIndex];

                                        // 判斷欄位值是否為指定欄位
                                        if (columns.Contains(cell.InnerText.Trim()))
                                        {
                                            // 獲取指定欄位的位置
                                            var columnPosition = Array.IndexOf(columns, cell.InnerText.Trim()) + 1;


                                            // 根據指定欄位的位置進行取值
                                            if (cell.InnerText.Trim() == "Delivery date")
                                            {
                                                // 從當前特定欄位的右方欄位取值
                                                var columnIndex = cellIndex + 1;
                                                if (columnIndex < cells.Count)
                                                {
                                                    var valueCell = cells[columnIndex];
                                                    var value = valueCell.InnerText.Trim();
                                                    worksheet.Cells[rowNumber - 1, columnPosition].Value = value;
                                                }
                                            }
                                            else if (cell.InnerText.Trim() == "Currency:")
                                            {

                                                // 從當前特定欄位的右方欄位取值
                                                var columnIndex = cellIndex + 1;
                                                if (columnIndex < cells.Count)
                                                {
                                                    var valueCell = cells[columnIndex];
                                                    var value = valueCell.InnerText.Trim();
                                                    worksheet.Cells[rowNumber - 1, columnPosition].Value = value;
                                                }

                                            }
                                            else if (cell.InnerText.Trim() == "Item")
                                            {

                                                // 從當前特定欄位的下方欄位取值
                                                var rowIndexBelow = rowIndex + 1;
                                                if (rowIndexBelow < rows.Count)
                                                {
                                                    var valueRow = rows[rowIndexBelow];
                                                    var valueCell = valueRow.SelectNodes("th|td")[cellIndex];
                                                    var value = valueCell.InnerText.Trim();
                                                    worksheet.Cells[rowNumber, columnPosition].Value = value;

                                                    rowNumber++;
                                                }


                                            }
                                            else if (cell.InnerText.Trim() == "PO Number")
                                            {
                                                // 從當前特定欄位的下方欄位取值
                                                var rowIndexBelow = rowIndex + 1;
                                                if (rowIndexBelow < rows.Count)
                                                {
                                                    var valueRow = rows[rowIndexBelow];
                                                    var valueCell = valueRow.SelectNodes("th|td")[cellIndex];
                                                    var value = valueCell.InnerText.Trim();
                                                    worksheet.Cells[rowNumber, columnPosition].Value = value;
                                                    worksheet.Cells[rowNumber + 1, columnPosition].Value = value;
                                                    worksheet.Cells[rowNumber + 2, columnPosition].Value = value;



                                                }

                                            }
                                            else if (cell.InnerText.Trim() == "Contact Person")
                                            {
                                                // 從當前特定欄位的下方欄位取值
                                                var rowIndexBelow = rowIndex + 1;
                                                if (rowIndexBelow < rows.Count)
                                                {
                                                    var valueRow = rows[rowIndexBelow];
                                                    var valueCell = valueRow.SelectNodes("th|td")[cellIndex - 1];
                                                    var value = valueCell.InnerText.Trim();
                                                    worksheet.Cells[rowNumber, columnPosition].Value = value;
                                                    worksheet.Cells[rowNumber + 1, columnPosition].Value = value;
                                                    worksheet.Cells[rowNumber + 2, columnPosition].Value = value;
                                                    worksheet.Cells[rowNumber, columnPosition + 1].Value = "USD";
                                                    worksheet.Cells[rowNumber + 1, columnPosition + 1].Value = "USD";
                                                    worksheet.Cells[rowNumber + 2, columnPosition + 1].Value = "USD";



                                                }

                                            }
                                            else
                                            {
                                                // 從下方欄位取值
                                                var rowIndexBelow = rowIndex + 1;
                                                if (rowIndexBelow < rows.Count)
                                                {
                                                    var valueRow = rows[rowIndexBelow];
                                                    var valueCell = valueRow.SelectNodes("th|td")[cellIndex];
                                                    var value = valueCell.InnerText.Trim();
                                                    worksheet.Cells[rowNumber - 1, columnPosition].Value = value;
                                                }
                                            }
                                        }
                                    }


                                }
                            }
                        }
                    }

                    // 儲存 Excel 文件
                    package.SaveAs(new FileInfo(outputFilePath));
                }

                Console.WriteLine("Excel 文件已生成。");
            }
            else
            {
                Console.WriteLine("郵件的正文不是 HTML 格式。");
            }

            Console.ReadLine();
        }
    }
}