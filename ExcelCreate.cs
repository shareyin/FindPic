using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Data;
using Aspose.Cells;

namespace Findpic
{
    
    /// <summary>
    ///OutFileDao ��ժҪ˵��
    /// </summary>
    public class ExcelCreate
    {
        public ExcelCreate()
        {
            //
            //TODO: �ڴ˴���ӹ��캯���߼�
            //
        }
        /// <summary>
        /// �������ݵ�����
        /// </summary>
        /// <param name="dt">Ҫ����������</param>
        /// <param name="tableName">������</param>
        /// <param name="path">����·��</param>
        public void OutFileToDisk(DataTable dt, string tableName, string path)
        {
            //try
            //{
                Workbook workbook = new Workbook(); //������
                Worksheet sheet = workbook.Worksheets[0]; //������
                Cells cells = sheet.Cells;//��Ԫ��


                //Ϊ����������ʽ    
                Style styleTitle = workbook.Styles[workbook.Styles.Add()];//������ʽ
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;//���־���
                styleTitle.Font.Name = "����";//��������
                styleTitle.Font.Size = 18;//���ִ�С
                styleTitle.Font.IsBold = true;//����

                //��ʽ2
                Style style2 = workbook.Styles[workbook.Styles.Add()];//������ʽ
                style2.HorizontalAlignment = TextAlignmentType.Center;//���־���
                style2.Font.Name = "����";//��������
                style2.Font.Size = 10;//���ִ�С
                style2.Font.IsBold = true;//����
                style2.IsTextWrapped = true;//��Ԫ�������Զ�����
                style2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                style2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                //��ʽ3
                Style style3 = workbook.Styles[workbook.Styles.Add()];//������ʽ
                style3.HorizontalAlignment = TextAlignmentType.Center;//���־���
                style3.Font.Name = "����";//��������
                style3.Font.Size = 9;//���ִ�С
                style3.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style3.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                style3.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style3.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                int Colnum = dt.Columns.Count;//�������
                int Rownum = dt.Rows.Count;//�������

                //������1 ������   
                //cells.Merge(0, 0, 1, Colnum);//�ϲ���Ԫ��
                //cells[0, 0].PutValue(tableName);//��д����
                //cells[0, 0].SetStyle(styleTitle);
                //cells.SetRowHeight(0, 38);

                //������2 ������
                for (int i = 0; i < Colnum; i++)
                {
                    cells[0, i].PutValue(dt.Columns[i].ColumnName);
                    //cells[0, i].SetStyle(style2);
                    //if (i == 4)
                    //{
                    //    cells.SetColumnWidth(i, 30);
                    //}
                    if (i == 6)
                    {
                        cells.SetColumnWidth(i, 30);
                    }
                    cells.SetRowHeight(0, 28);
                }

                //����������
                for (int i = 0; i < Rownum-1; i++)
                {

                    for (int k = 0; k < Colnum; k++)
                    {
                        //if (k == 1)
                        //{
                            cells[1 + i, k].PutValue(dt.Rows[i][k]);
                        //}
                        //else
                        //{
                        //    cells[2 + i, k].PutValue(Convert.ToInt32(dt.Rows[i][k].ToString()));
                        //}

                        //cells[1 + i, k].SetStyle(style3);

                    }
                    cells.SetRowHeight(i, 18);
                }

                workbook.Save(path);
        
        }


        public MemoryStream OutFileToStream(DataTable dt, string tableName)
        {
            Workbook workbook = new Workbook(); //������
            Worksheet sheet = workbook.Worksheets[0]; //������
            Cells cells = sheet.Cells;//��Ԫ��

            //Ϊ����������ʽ    
            Style styleTitle = workbook.Styles[workbook.Styles.Add()];//������ʽ
            styleTitle.HorizontalAlignment = TextAlignmentType.Center;//���־���
            styleTitle.Font.Name = "����";//��������
            styleTitle.Font.Size = 18;//���ִ�С
            styleTitle.Font.IsBold = true;//����

            //��ʽ2
            Style style2 = workbook.Styles[workbook.Styles.Add()];//������ʽ
            style2.HorizontalAlignment = TextAlignmentType.Center;//���־���
            style2.Font.Name = "����";//��������
            style2.Font.Size = 12;//���ִ�С
            style2.Font.IsBold = true;//����
            style2.IsTextWrapped = true;//��Ԫ�������Զ�����
            style2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            //��ʽ3
            Style style3 = workbook.Styles[workbook.Styles.Add()];//������ʽ
            style3.HorizontalAlignment = TextAlignmentType.Center;//���־���
            style3.Font.Name = "����";//��������
            style3.Font.Size = 9;//���ִ�С
            style3.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style3.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style3.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style3.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            int Colnum = dt.Columns.Count;//�������
            int Rownum = dt.Rows.Count;//�������

            //////������1 ������   
            //cells.Merge(0, 0, 1, Colnum);//�ϲ���Ԫ��
            //cells[0, 0].PutValue(tableName);//��д����
            //cells[0, 0].SetStyle(styleTitle);
            //cells.SetRowHeight(0, 38);

            ////������2 ������
            //for (int i = 0; i < Colnum; i++)
            //{
            //    cells[1, i].PutValue(dt.Columns[i].ColumnName);
            //    cells[1, i].SetStyle(style2);
            //    cells.SetRowHeight(1, 30);
            //}

            //����������
            for (int i = 0; i < Rownum; i++)
            {
                for (int k = 0; k < Colnum; k++)
                {
                    cells[2 + i, k].PutValue(dt.Rows[i][k].ToString());
                    cells[2 + i, k].SetStyle(style3);
                }
                cells.SetRowHeight(2 + i, 18);
            }

            MemoryStream ms = workbook.SaveToStream();
            return ms;
        }

    }
}
