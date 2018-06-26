using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace fbdetonator
{
    class FBDonateRules
    {
        public enum Col { Donation = 3, Date = 11, Title = 16, Source };

        private int mPageDonateBtnCnt;
        private int mPostDonateBtnCnt;
        private int mFundraiserCnt;

        public FBDonateRules()
        {
            mPageDonateBtnCnt = 0;
            mPostDonateBtnCnt = 0;
            mFundraiserCnt = 0;
        }

        // Define getters/setters for member vars

        // Define biz logic methods
        // - GetFundraiserCnt
        // - GetPageDonateCnt
        // - GetPostDonateCnt

        public bool IsPostDonation(string value)
        {
            mPostDonateBtnCnt++;
            return value.Equals("donate_button_user_posts");
        }

        public bool IsPageDonation(string value)
        {
            mPageDonateBtnCnt++;
            return value.Equals("donate_button_charity_page");
        }

        public bool IsFundraiserDonation(string value)
        {
            mFundraiserCnt++;
            return value.Equals("fundraiser");
        }
    }
}
