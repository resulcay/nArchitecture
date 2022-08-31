using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<string> dataList = new List<string>();

            int mainCounter = 0;
            int tempCounter = 0;

            string filePath = @"C:\Users\Resul\Desktop\5163.xlsx";
            string sheet = "Sheet1";
            string columnFrom = "L";
            string columnTo = "L";
            string cellMustStartWith = "202";
            string cellMustContain = "101";
            string cellMustEndWith = "204";

            using (var excelWorkbook = new XLWorkbook(filePath))
            {
                var Ws = excelWorkbook.Worksheet(sheet);
                dataList = Ws.Range(columnFrom+":"+columnTo)
                    .CellsUsed()       
                    .Select(c => c.Value.ToString()) 
                    .ToList();
            }
            dataList.RemoveAt(0);

            for (int i = 0; i < dataList.Count; i++)
            {
                if (dataList[i] == cellMustStartWith)
                {
                    for (int m = i + 1; m < dataList.Count && dataList[m] != cellMustStartWith; m++)
                    {
                        if (dataList[m] == cellMustContain)
                        {
                            tempCounter++;
                        }

                        if (dataList[m]==cellMustEndWith)
                        {
                            mainCounter += tempCounter;
                            tempCounter = 0;
                            break;
                        }
                    }
                    tempCounter = 0;
                }
            }
            Console.WriteLine(mainCounter);
        }
    }
} 