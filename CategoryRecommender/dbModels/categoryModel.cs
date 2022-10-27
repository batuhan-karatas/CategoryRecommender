using System;
using System.Collections.Generic;
using excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace CategoryRecommender.dbModels
{
    public class categoryModel
    {
        public int id { get; set; }

        public string title { get; set; }

        public string tags { get; set; }


        public List<categoryModel> ReadFromCategory(String dataPath)
        {
            List<categoryModel> resultList = new List<categoryModel>();



            excel.Application app = new excel.Application();
            excel.Workbook workbook = app.Workbooks.Open(dataPath);
            excel.Worksheet worksheet = (Worksheet)workbook.Sheets[1];

            int excelRowRange = worksheet.UsedRange.Rows.Count;
            int excelColumnRange = worksheet.UsedRange.Columns.Count;



            for (int row = 2; row <= 200; row++)
            {
                categoryModel model = new categoryModel();

                for (int column = 1; column <= 15; column++)
                {
                    excel.Range col = (excel.Range)worksheet.Cells[1, column];
                    excel.Range range = (excel.Range)worksheet.Cells[row, column];
                    if (col.Value.ToString() == "ID")
                    {
                        if (range.Value != null)
                        {

                            int id = Convert.ToInt32(range.Value);
                            model.id = id;
                        }
                        else
                        {
                            goto LoopEnd;
                        }
                    }
                    else if (col.Value.ToString() == "Title")
                    {

                        model.title = range.Value.ToString();
                    }
                    else if (col.Value.ToString() == "TAGS")
                    {
                        if (range.Value != null)
                        {
                            model.tags = range.Value.ToString();

                        }

                    }


                }
                resultList.Add(model);

            }
        LoopEnd:
            workbook.Close();
            app.Quit();
            return resultList;


        }
    }
}
