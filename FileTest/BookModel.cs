using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
 
namespace FileTest
{
    public class BookModel
    {
        /// <summary>
        /// 所对应的课程类型
        /// </summary>
        private string bookType;
 
        public string BookType
        {
            get { return bookType; }
            set { bookType = value; }
        }
 
        /// <summary>
        /// 书所对应的ISBN号
        /// </summary>
        private string bookISBN;
 
        public string BookISBN
        {
            get { return bookISBN; }
            set { bookISBN = value; }
        }
 
        /// <summary>
        /// 书名
        /// </summary>
        private string bookName;
 
        public string BookName
        {
            get { return bookName; }
            set { bookName = value; }
        }
 
        /// <summary>
        /// 作者
        /// </summary>
        private string bookAuthor;
 
        public string BookAuthor
        {
            get { return bookAuthor; }
            set { bookAuthor = value; }
        }
 
        /// <summary>
        /// 价格
        /// </summary>
        private double bookPrice;
 
        public double BookPrice
        {
            get { return bookPrice; }
            set { bookPrice = value; }
        }
    }
}
