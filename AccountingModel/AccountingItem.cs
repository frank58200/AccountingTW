using System;
namespace AccountingModel
{
    /// <summary>
    /// 台灣 經濟部 商業會計科目表
    /// </summary>
    public class AccountingItem
    {
        public AccountingItem()
        {

        }

        public string Item1stId { get; set; }
        public string Item2ndId { get; set; }
        public string Item3rdId { get; set; }
        public string Item4thId { get; set; }

        public string ItemId { get; set; }

        public string AccountNameCh { get; set; }
        public string AccountNameEn { get; set; }

        public string DescriptionCh { get; set; }
        public string DescriptionEn { get; set; }

    }
}
