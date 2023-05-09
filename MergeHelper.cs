using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;

namespace WordHelperLibrary
{
    /// <summary>
    /// 合并帮助类
    /// </summary>
    public static class MergeHelper
    {
        public static string Merge(List<string> inputFilenameList, string outputFileName)
        {
            try
            {
                var sourceList = new List<Source>();
                foreach (var inputDoc in inputFilenameList)
                {
                    var fs = new FileStream(inputDoc, FileMode.Open);
                    var tempms = new MemoryStream();
                    fs.CopyTo(tempms);
                    var source = new Source(new WmlDocument(fs.Length.ToString(), tempms), true);
                    sourceList.Add(source);
                    tempms.Close();
                    fs.Close();
                }
                var mergedDoc = DocumentBuilder.BuildDocument(sourceList);
                mergedDoc.SaveAs(outputFileName);
                return "";
            }
            catch (Exception ex)
            {
                return $"合并失败，原因：{ex.Message}";
            }
        }
    }
}
