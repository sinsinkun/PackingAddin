using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPacking
{
    public partial class ArrayRibbon
    {

        Excel.Application xlApp = null;
        Excel.Workbook xlWorkbook = null;
        Excel.Worksheet xlWorksheet = null;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //initialization
        }

        private void Btn_Pack_Click(object sender, RibbonControlEventArgs e)
        {
            //Connect to excel
            xlApp = Globals.ThisAddIn.Application;
            xlWorkbook = xlApp.ActiveWorkbook;
            xlWorksheet = xlWorkbook.ActiveSheet;

            /* 0. check if excel sheet is correct one */
            if (xlWorkbook.Name != "Sheetstock Excel.xlsx" || xlWorksheet.Name != "Calculator")
            {
                DialogResult yesNo = MessageBox.Show("The active Excel sheet does not appear to be the official Sheetstock Calculator Excel sheet. This program may not work as intended. \n\nProceed Anyway?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (yesNo == DialogResult.No) return;
            }
            Debug.Print("Workbook: {0} -- Worksheet: {1}", xlWorkbook.Name, xlWorksheet.Name);

            Rect bounds = new Rect();
            Rect topLeftCoord = new Rect { x = 0, y = 0 };
            List<Rect> rectList = new List<Rect>();
            double tempx, tempy;
            int tempq, row = 2;
            Color tempc;
            bool b = true;
            string savePath = xlWorkbook.Path + "\\Sheets\\";

            /* 1. load in all data */
            #region Loading in data for sorting and packing
            //save bounds size
            tempx = xlWorksheet.Cells[3, 9].Value;
            tempy = xlWorksheet.Cells[4, 9].Value;
            if (tempx > tempy) bounds = new Rect { x = tempx, y = tempy };
            else bounds = new Rect { x = tempy, y = tempx };
            Debug.Print("Bounds: {0}, {1}", bounds.x, bounds.y);

            //check if first row is filled
            if (xlWorksheet.Cells[row, 1].Value == null && xlWorksheet.Cells[row, 2].Value == null && xlWorksheet.Cells[row, 3].Value == null)
            {
                MessageBox.Show("First row is not filled, or there is no input. Please make sure all data is filled properly.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //fill rectList
            while (b)
            {
                if (xlWorksheet.Cells[row, 1].Value != null && xlWorksheet.Cells[row, 2].Value != null && xlWorksheet.Cells[row, 3].Value != null && xlWorksheet.Cells[row, 4].Value != null)
                {
                    tempx = xlWorksheet.Cells[row, 1].Value;
                    tempy = xlWorksheet.Cells[row, 2].Value;
                    tempq = (int)xlWorksheet.Cells[row, 3].Value;
                    tempc = Color.FromName(xlWorksheet.Cells[row, 4].Value.ToString());
                    Debug.Print("Saving to list: x: {0}, y: {1}, quantity: {2}, color: {3}", tempx, tempy, tempq, tempc.ToKnownColor());
                    rectList.Add(new Rect { x = tempx, y = tempy, quant = tempq, penColor = tempc });
                    row++;
                }
                else if (xlWorksheet.Cells[row, 1].Value == null && xlWorksheet.Cells[row, 2].Value == null && xlWorksheet.Cells[row, 3].Value == null)
                {
                    b = false;
                }
                else
                {
                    MessageBox.Show("There is an incomplete row of data. Please fix.");
                    return;
                }
            }

            #endregion

            MessageBox.Show("Data imported. \n\nOutput will be printed to: " + savePath + "\n\nPlease note that any existing output files in this directory will be overwritten.",
                            "Import Complete", MessageBoxButtons.OK);
            Debug.Print("Number of objects: {0}", rectList.Count);

            //convert rectlist to rect array
            Rect[] rect = new Rect[rectList.Count];
            rect = rectList.ToArray();

            //create graphic
            Bitmap bmp = new Bitmap(Convert.ToInt32(bounds.x * 10), Convert.ToInt32(bounds.y * 10));
            Graphics g = Graphics.FromImage(bmp);
            g.FillRectangle(Brushes.White, new Rectangle(0, 0, Convert.ToInt32(bounds.x * 10), Convert.ToInt32(bounds.y * 10)));

            //check if output file path exists
            try
            {
                if (!System.IO.Directory.Exists(savePath)) System.IO.Directory.CreateDirectory(savePath);
            }
            catch (Exception)
            {
                MessageBox.Show("Cannot create output directory", "Folder Creation Error", MessageBoxButtons.OK);
                return;
            }

            /* 2. sort data */
            #region sort from largest to smallest
            b = true;
            double tempD;
            Rect tempR = new Rect();
            Rect absMin = new Rect { x = bounds.x, y = bounds.y };
            Rect absMax = new Rect { x = 0, y = 0 };

            // sort by longest edge (x)
            for (int i = 0; i < rect.Length; i++)
            {
                // align so longest edge along x axis
                if (rect[i].x <= rect[i].y)
                {
                    tempD = rect[i].x;
                    rect[i].x = rect[i].y;
                    rect[i].y = tempD;
                }
                if (rect[i].x < absMin.x) absMin = rect[i];
                if (rect[i].x >= absMax.x) absMax = rect[i];
            }
            while (b)
            {
                for (int i = 0; i < rect.Length - 1; i++)
                {
                    if (rect[i].x <= rect[i + 1].x)
                    {
                        tempR = rect[i];
                        rect[i] = rect[i + 1];
                        rect[i + 1] = tempR;
                    }
                }
                if (rect[0] == absMax && rect[rect.Length - 1] == absMin) b = false;
            }
            #endregion

            /* 3. packing algorithm */
            #region piece by piece packing method
            //find total number of objects
            int quantAll = 0;
            for (int i = 0; i < rect.Length; i++) quantAll += rect[i].quant;
            Debug.Print("total # of elements: {0} \n", quantAll);

            //place 1 by 1
            Rect remainingRowArea = new Rect { x = bounds.x, y = bounds.y };
            Rect smallestRemainingRect = new Rect();
            double nextRowx = 0;
            int sheetNum = 1;
            int[] counter = new int[rect.Length];
            foreach (int i in counter) counter[i] = 0;
            bool endofSheet = false;

            for (int i = 0; i < quantAll; i++)
            {
                //If no more rectangles can be drawn
                if (i < quantAll - 1 && endofSheet)
                {
                    string message = "No more rectangles can be fitted onto this sheet, please move to next sheet. \n\n";
                    for (int j = 0; j < rect.Length; j++)
                    {
                        message += "Remaining # of " + rect[j].penColor.ToKnownColor() + " rectangles: " + (rect[j].quant - counter[j]) + "\n";
                    }
                    Debug.Print(message);
                    // save bmp, reset and start again
                    bmp.Save(savePath + "Sheet-" + sheetNum + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                    g.FillRectangle(Brushes.White, new Rectangle(0, 0, Convert.ToInt32(bounds.x * 10), Convert.ToInt32(bounds.y * 10)));
                    endofSheet = false;
                    sheetNum++;
                    i--;
                    topLeftCoord.x = 0;
                    topLeftCoord.y = 0;
                    remainingRowArea.x = bounds.x;
                    remainingRowArea.y = bounds.y;
                }
                //find smallest remaining piece
                for (int j = rect.Length - 1; j > 0; j--)
                {
                    if (counter[j] < rect[j].quant)
                    {
                        smallestRemainingRect = rect[j];
                        break;
                    }
                }
                //if remaining area is smaller than smallest remaining piece
                if (remainingRowArea.y < smallestRemainingRect.x && remainingRowArea.y < smallestRemainingRect.y)
                {
                    // goto next row
                    topLeftCoord.y = 0;
                    topLeftCoord.x += nextRowx;
                    remainingRowArea.y = bounds.y;
                    remainingRowArea.x -= nextRowx;
                    nextRowx = 0;
                    Debug.Print("Going to next row - remaining area: {0} {1}", remainingRowArea.x, remainingRowArea.y);
                }
                //go through pre sorted rectangles (longest x to shortest)
                for (int j = 0; j < rect.Length; j++)
                {
                    //if there is enough space to fit 2 of this rectangle in current row, draw until no more fits
                    if (counter[j] < rect[j].quant && rect[j].y <= remainingRowArea.y && rect[j].x * 2 <= nextRowx)
                    {
                        Debug.Print("Instance {0} - At least 2 " + rect[j].penColor.ToKnownColor() + " Rectangles can fit here", i);
                        double originalx = topLeftCoord.x;
                        double remainingRowx = nextRowx;
                        i--;
                        while (counter[j] < rect[j].quant && rect[j].y <= remainingRowArea.y && rect[j].x <= remainingRowx)
                        {
                            g.DrawRectangle(rect[j].pen, Convert.ToInt32(topLeftCoord.x * 10), Convert.ToInt32(topLeftCoord.y * 10), Convert.ToInt32(rect[j].x * 10), Convert.ToInt32(rect[j].y * 10));
                            topLeftCoord.x += rect[j].x;
                            remainingRowx -= rect[j].x;
                            counter[j]++;
                            i++;
                            Debug.Print("Instance {0} - Another " + rect[j].penColor.ToKnownColor() + " Rectangle was drawn in parallel, #{1} \n", i, counter[j]);
                        }
                        topLeftCoord.x = originalx;
                        topLeftCoord.y += rect[j].y;
                        remainingRowArea.y -= rect[j].y;
                        break;
                    }
                    //draws ONE rectangle if there is remaining amount & fits into remaining area (biggest to smallest)
                    else if (counter[j] < rect[j].quant && rect[j].y <= remainingRowArea.y && rect[j].x <= remainingRowArea.x)
                    {
                        g.DrawRectangle(rect[j].pen, Convert.ToInt32(topLeftCoord.x * 10), Convert.ToInt32(topLeftCoord.y * 10), Convert.ToInt32(rect[j].x * 10), Convert.ToInt32(rect[j].y * 10));
                        topLeftCoord.y += rect[j].y;
                        remainingRowArea.y -= rect[j].y;
                        counter[j]++;
                        if (rect[j].x > nextRowx) nextRowx = rect[j].x;
                        Debug.Print("Instance {0} - Drew new " + rect[j].penColor.ToKnownColor() + " rectangle, #{1} \n", i, counter[j]);
                        break;
                    }
                    else
                    {
                        Debug.Print("Instance {0} - " + rect[j].penColor.ToKnownColor() + " Rectangle did not pass conditions", i);

                        //See if reached end of sheet
                        if (j == rect.Length - 1) endofSheet = true;
                    }
                }
                if (i == quantAll - 1) Debug.Print("Completed. Final instance: {0}", i);
            }

            //print and save last sheet
            bmp.Save(savePath + "Sheet-" + sheetNum + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            MessageBox.Show("Total number of sheets necessary: " + sheetNum, "Packing completed", MessageBoxButtons.OK);
            #endregion

            /* 4. clean up */
            xlWorksheet.Cells[7, 9].Value = sheetNum;
            topLeftCoord.x = 0;
            topLeftCoord.y = 0;
            rectList.Clear();

        }
    }

    public class Rect
    {
        public double x { get; set; }
        public double y { get; set; }
        public int quant { get; set; }
        public Color penColor { get; set; }
        public double area
        {
            get { return x * y; }
        }
        public Pen pen
        {
            get { return new Pen(penColor, 2); }
        }

    }

}
