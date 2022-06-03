using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Advanced.Common
{
    /// <summary>
    /// Excel帮助类库
    /// </summary>
    public class ExcelOperationHelper
    {
        public static IWorkbook CreateExcelWorkBook(bool isxls = true)
        {
            dynamic _Workbook;
            if (isxls) 
            {
                //默认写xls格式
                _Workbook = new HSSFWorkbook();
            }
            else
            {
                //否则写xlsx格式
                _Workbook = new XSSFWorkbook();
            }

            ISheet sheet1 = _Workbook.CreateSheet("Sheet1");

            //创建第一行
            {
                IRow head = sheet1.CreateRow(0);
                ICell cell = head.CreateCell(0);
                cell.SetCellValue("学生姓名");

                ICell cell1 = head.CreateCell(1);
                cell1.SetCellValue("数学成绩");

                ICell cell2 = head.CreateCell(2);
                cell2.SetCellValue("语文成绩");
            }

            //创建第二行
            {
                IRow head = sheet1.CreateRow(1);
                ICell cell = head.CreateCell(0);
                cell.SetCellValue("Jeffrey");

                ICell cell1 = head.CreateCell(1);
                cell1.SetCellValue("100");

                ICell cell2 = head.CreateCell(2);
                cell2.SetCellValue("95");
            }

            return _Workbook;
        }

        public static IWorkbook DataToHSSFWorkbook(List<ExcelDataResource> dataResources)
        {
            HSSFWorkbook _Workbook = new HSSFWorkbook();

            if (dataResources.Equals(null) || dataResources.Count.Equals(0))
            {
                return _Workbook;
            }
            //每循环一次，就生成一个Sheet页出来
            foreach(var sheetResource in dataResources)
            {
                if(sheetResource.SheetDataResource.Equals(null) || sheetResource.SheetDataResource.Count.Equals(0))
                {
                    break;
                }

                //创建一个页签
                ISheet sheet = _Workbook.CreateSheet(sheetResource.SheetName);
                //确定当前这一页有多少列--取决于保存当前Sheet页数据的实体结构中的 标记特性
                object obj = sheetResource.SheetDataResource[0];

                //获取需要导出的所有列
                Type type = obj.GetType();
                List<PropertyInfo> propList = type.GetProperties().Where(r => r.IsDefined(typeof(TitleAttribute), true)).ToList();


                //确定表头在哪一行生成
                int titleIndex = sheetResource.HeadIndex >= 0 ? sheetResource.HeadIndex : 0;
                //基于当前Sheet页创建表头
                IRow titleRow = sheet.CreateRow(titleIndex);


                //给表头创建单元格，并用特性中属性名填充值
                for(int i = 0; i < propList.Count(); i++)
                {
                  TitleAttribute propertyAttribute =  propList[i].GetCustomAttribute<TitleAttribute>();

                  ICell cell = titleRow.CreateCell(i);
                  cell.SetCellValue(propertyAttribute.Title);
                }

                //生成数据
                for (int i = 0; i < sheetResource.SheetDataResource.Count() ; i++)
                {
                    IRow row = sheet.CreateRow(i+ titleIndex);
                    object objInstance = sheetResource.SheetDataResource[i];

                    for(int j = 0; j < propList.Count; j++)
                    {
                        ICell cell = titleRow.CreateCell(j);
                        cell.SetCellValue(propList[j].GetValue(objInstance).ToString());
                    }
                }
            }

            return _Workbook;
        }
    }
}
