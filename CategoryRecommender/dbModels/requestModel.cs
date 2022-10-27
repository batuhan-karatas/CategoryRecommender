using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using excel = Microsoft.Office.Interop.Excel;

namespace CategoryRecommender.dbModels
{
    public class requestModel
    {
        public int id { get; set; }

        public string title { get; set; }


        public int status { get; set; }

        public int catId { get; set; }

        public string text { get; set; }



        public List<requestModel> ReadFromRequest(String dataPath)
        {
            List<requestModel> resultList = new List<requestModel>();



            excel.Application app = new excel.Application();
            excel.Workbook workbook = app.Workbooks.Open(dataPath);
            excel.Worksheet worksheet = (Worksheet)workbook.Sheets[1];

            int excelRowRange = worksheet.UsedRange.Rows.Count;
            int excelColumnRange = worksheet.UsedRange.Columns.Count;


            for (int row = 2; row <= excelRowRange; row++)
            {
                requestModel model = new requestModel();
                
                for (int column = 1; column <= excelColumnRange; column++)
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
                   
                    else if (col.Value.ToString() == "STATUS")
                    {
                        int status = Convert.ToInt32(range.Value);
                        if (status != 200 || status != 201)
                        {
                            model.status = status;
                        }
                    }
                    else if (col.Value.ToString() == "CAT_ID")
                    {
                        if (range.Value != null)
                        {

                            int id1 = Convert.ToInt32(range.Value);
                            model.catId = id1;
                        }
                        
                        

                    }
                    else if (col.Value.ToString() == "TEXT")
                    {
                        model.text = range.Value.ToString();
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
