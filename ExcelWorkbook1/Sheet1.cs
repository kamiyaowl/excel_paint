using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Threading.Tasks;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ExcelWorkbook1
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            var dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.Columns.ColumnWidth = 1;
                this.Rows.RowHeight = 10;


                var bitmap = new Bitmap(dialog.FileName);
                var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                                                        ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                int width = bitmapData.Width;
                int height = bitmapData.Height;

                byte[] bytes = new byte[bitmapData.Stride * bitmapData.Height];
                Marshal.Copy(bitmapData.Scan0, bytes, 0, bytes.Length);
                int index = 0;
                int nResidual = bitmapData.Stride - width * 3;

                for (int i = 0; i < height; i++)
                {
                    for (int j = 0; j < width; j++)
                    {
                        int r = bytes[index + 2];
                        int g = bytes[index + 1];
                        int b = bytes[index];
                        var cell = this.Cells[j + 1][i + 1];
                        cell.Interior.Color = Color.FromArgb(r, g, b);

                        index += 3;
                    }
                    index += nResidual;
                }
                bitmap.UnlockBits(bitmapData);

            }

        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO デザイナーで生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet1_Startup);
            this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
        }

        #endregion

    }
}
