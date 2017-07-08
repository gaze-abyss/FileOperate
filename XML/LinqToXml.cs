using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace FileOperate.XML
{
    //1: xn 代表一个结点
    //2: xn.Name;//这个结点的名称
    //3: xn.Value;//这个结点的值
    //4: xn.ChildNodes;//这个结点的所有子结点
    //5: xn.ParentNode;//这个结点的父结点
    public class LinqToXml
    {
        /// <summary>
        /// 读取某一个节点下的所有数据
        /// </summary>
        /// <param name="path"></param>
        /// <param name="xnName"></param>
        public void ReadAllByXnName(string path,string xnName)
        {
            XElement xe = XElement.Load(path);
            IEnumerable<XElement> elements = from ele in xe.Elements("book") select ele;
        }

        public void InsertRecord(string path)
        {

        }

        //private void showInfoByElements(IEnumerable<XElement> elements)
        //{
        //    List<BookModel> modelList = new List<BookModel>();
        //    foreach (var ele in elements)
        //    {
        //        BookModel model = new BookModel();
        //        model.BookAuthor = ele.Element("author").Value;
        //        model.BookName = ele.Element("title").Value;
        //        model.BookPrice = Convert.ToDouble(ele.Element("price").Value);
        //        model.BookISBN=ele.Attribute("ISBN").Value;
        //        model.BookType=ele.Attribute("Type").Value;
                
        //        modelList.Add(model);
        //    }
        //    dgvBookInfo.DataSource = modelList;
        //}
    }
}
