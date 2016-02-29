using System;

namespace ExcelToolsGlobalNamingArea
{
    public class ExcelToolsVM
    {
        #region Properties
        public string SheetNm { get; set; }
        public string SheetNo1 { get; set; }
        public string SheetNo2 { get; set; }
        public int No1 { get; private set; }
        public int No2 { get; private set; }
        public int MaxLoopCnt { get; set; }
        #endregion

        #region Required Items Validation
        internal bool NotNullValidation()
        {
            return this.IsNotNullOrEmpty();
        }

        private bool IsNotNullOrEmpty()
        {
            return (!string.IsNullOrEmpty(this.SheetNo1)
                && !string.IsNullOrEmpty(this.SheetNo2));
        }
        #endregion

        #region Single Item Validation
        internal bool SingleValidation()
        {
            return this.IsNumeric()
                && this.IsNotOverValue();
        }

        private bool IsNumeric()
        {
            int tmpNo1, tmpNo2;

            if (int.TryParse(this.SheetNo1.ToString(), out tmpNo1)
                && int.TryParse(this.SheetNo2.ToString(), out tmpNo2))
            {
                this.No1 = tmpNo1;
                this.No2 = tmpNo2;
                return true;
            }

            return false;
        }
    
        private bool IsNotOverValue()
        {
            return (this.No1 < 101 || this.No2 < 101);
        }
        #endregion

        #region Multi Items Validation
        internal bool MultiValidation()
        {
            return this.IsRangeCorrect();
        }

        private bool IsRangeCorrect()
        {
            return (this.No1 <= this.No2);
        }

        internal bool IsEnoughSheets()
        {
            return this.IsMaxLoopCntEnough();
        }

        private bool IsMaxLoopCntEnough()
        {
            this.MaxLoopCnt = this.No2 - this.No1;
            return (this.MaxLoopCnt > 1);
        }
        #endregion
    }
}
