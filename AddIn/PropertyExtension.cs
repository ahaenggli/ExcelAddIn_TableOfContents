using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn_TableOfContents
{
    class PropertyExtension
    {

        //' get id of custom property by name
        public static int getPropId(Excel.Worksheet ws, String propName)
        {
            int tmp;

            if (ws == null || String.IsNullOrWhiteSpace(propName))
            {
                return 0;
            }

            tmp = 0;
            foreach (Excel.CustomProperty xx in ws.CustomProperties)
            {
                tmp = tmp + 1;
                if (xx.Name.ToLower().Equals(propName.ToLower())) return tmp;
            }

            return 0;
        }



        // rename and overwrite property
        public static void propertyRename(Excel.Workbook WB, String propNameOld, String propNameNew)
        {

            foreach (Excel.Worksheet ws in WB.Worksheets)
            {
                setProperty(ws, propNameNew, getProperty(ws, propNameOld));
                setProperty(ws, propNameOld, "");
            }
        }


        //' get value of custom property by name, default propName is "Tag"
        public static String getProperty(Excel.Worksheet ws, String propName)
        {
            int propId;

            if (ws == null || String.IsNullOrWhiteSpace(propName)) return null;


            propId = getPropId(ws, propName);

            if (propId > 0)
            {
                return ws.CustomProperties.Item[propId].Value;
            }
            return null;
        }

        //'set value of custom propery by name, default propName is "Tag"
        public static void setProperty(Excel.Worksheet ws, String propName, String propVal)
        {
            if (ws == null || String.IsNullOrWhiteSpace(propName)) return;

            int propId;
            propId = getPropId(ws, propName);

            //    'delete if exists
            if (propId > 0)
            {
                ws.CustomProperties.Item[propId].Delete();
            }

            //    ' Add metadata to worksheet.
            if (!String.IsNullOrWhiteSpace(propVal)) ws.CustomProperties.Add(propName, propVal);
        }



    }
}
