using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Data;

namespace RFQEventReceiver.Entities
{
    public class ColumnInfo
    {
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="t"></param>
        ///// <returns></returns>
        //public object[,] GetFields(Type t)
        //{

        //    FieldInfo[] oFields = t.GetFields();
        //    FieldInfo oField;
        //    Attribute[] attributes;
        //    int fieldLength = oFields.Length, i;
        //    object[,] StructureInfo = new object[fieldLength, 2];

        //    try
        //    {
        //        for (i = 0; i < fieldLength; i++)
        //        {
        //            oField = oFields[i];
        //            attributes = Attribute.GetCustomAttributes(oField, typeof(ColumnAttributes), false);
        //            StructureInfo[i, 0] = oField;
        //            StructureInfo[i, 1] = attributes;
        //        }

        //    }
        //    catch (Exception ex)
        //    {                
        //    }
        //    return StructureInfo;
        //}

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="StructureInfo"></param>
        ///// <param name="oRow"></param>
        //public void SetExcelFields(object[,] StructureInfo, DataRow oRow)
        //{
        //    FieldInfo oField;
        //    Attribute[] attributes;

        //    try
        //    {
        //        int upperBound = StructureInfo.GetUpperBound(0), i;

        //        for (i = 0; i <= upperBound; i++)
        //        {
        //            oField = (FieldInfo)StructureInfo[i, 0];
        //            attributes = (Attribute[])StructureInfo[i, 1];

        //            foreach (Attribute attr in attributes)
        //            {
        //                ColumnAttributes oColumnAttributeName = (ColumnAttributes)attr;
        //                if (oRow[oColumnAttributeName.ExcelColumnName] != System.DBNull.Value)
        //                {
        //                    oField.SetValue(this, oRow[oColumnAttributeName.ExcelColumnName]);
        //                }
        //                break;
        //            }
        //        }

        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }

        //    return;
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        public object[,] GetProperties(Type t)
        {
            PropertyInfo[] oProperties = t.GetProperties();
            PropertyInfo oProperty;
            Attribute[] attributes;
            int propertyLength = oProperties.Length, i;
            object[,] StructureInfo = new object[propertyLength, 2];

            try
            {
                for (i = 0; i < propertyLength; i++)
                {
                    oProperty = oProperties[i];
                    attributes = Attribute.GetCustomAttributes(oProperty, typeof(ColumnAttributes), false);
                    StructureInfo[i, 0] = oProperty;
                    StructureInfo[i, 1] = attributes;
                }
            }
            catch (Exception ex)
            {
            }
            return StructureInfo;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="StructureInfo"></param>
        /// <param name="oRow"></param>
        public void SetExcelProperties(object[,] StructureInfo, DataRow oRow)
        {
            PropertyInfo oProperty;
            Type oPropertyType;
            Attribute[] attributes;

            try
            {
                int upperBound = StructureInfo.GetUpperBound(0), i;

                for (i = 0; i <= upperBound; i++)
                {
                    oProperty = (PropertyInfo)StructureInfo[i, 0];
                    oPropertyType = oProperty.PropertyType;
                    attributes = (Attribute[])StructureInfo[i, 1];

                    foreach (Attribute attr in attributes)
                    {
                        ColumnAttributes oColumnAttributeName = (ColumnAttributes)attr;
                        string attrNm = oColumnAttributeName.ExcelColumnName;
                        try
                        {
                            if (attrNm != null && // only set this property if the attribute's ExcelColumnName is not null ...
                                oRow[attrNm] != System.DBNull.Value) // ... and the value from the row is not null
                            {
                                //switch (oPropertyType) // 
                                //{
                                //    case "System.DateTime":
                                //        oProperty.SetValue(this, ExcelDocumentUtil.ReadExcelDateTimeValue(oRow[attrNm]), null);
                                //        break;
                                //    case "System.Int32":
                                //        oProperty.SetValue(this, Int32.Parse(oRow[attrNm].ToString()), null);
                                //        break;
                                //    default:
                                //        oProperty.SetValue(this, oRow[attrNm], null);
                                //        break;
                                //}

                                if (oPropertyType == typeof(DateTime))
                                {
                                    oProperty.SetValue(this, ExcelDocumentUtil.ReadExcelDateTimeValue(oRow[attrNm]), null);
                                }
                                else if (oPropertyType == typeof(Int32))
                                {
                                    oProperty.SetValue(this, Int32.Parse(oRow[attrNm].ToString()), null);
                                }
                                else
                                {
                                    oProperty.SetValue(this, oRow[attrNm], null);
                                }
                            }
                            break;
                        }
                        catch (ArgumentException argEx)
                        {
                            //if (argEx.Message.Contains("DateTime"))
                            //{
                            //    oProperty.SetValue(this, ExcelDocumentUtil.ReadExcelDateTimeValue(oRow[attrNm]), null);
                            //}
                            //else if (argEx.Message.Contains("Int32"))
                            //{
                            //    oProperty.SetValue(this, int.Parse(oRow[attrNm].ToString()), null);
                            //}
                            //else
                            //{

                            oProperty.SetValue(this, "Retrieve Error", null);

                            //}
                        }
                    } // end if 
                }

            }
            catch (Exception)
            {
                throw;
            }

            return;
        }
    }

    //[AttributeUsage(AttributeTargets.Field, AllowMultiple = true)]
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ColumnAttributes : System.Attribute
    {
        public string SPColumnName = "";
        public string ExcelColumnName = "";
        public bool IsHistoryData = false;

        public ColumnAttributes(string SharePointColumnName, string ExcelColumnName)
        {
            this.SPColumnName = SharePointColumnName;
            this.ExcelColumnName = ExcelColumnName;
        }

        public ColumnAttributes(string SharePointColumnName, string ExcelColumnName, bool IsHistData)
        {
            this.SPColumnName = SharePointColumnName;
            this.ExcelColumnName = ExcelColumnName;
            this.IsHistoryData = IsHistData;
        }
    }
}
