using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CS_KPMCreator 
{
    class WebControl
    {
        private void CreateOneTicket(Dictionary<string, string> dItem)
        {
            // open Browser

            // Log In

            // Action one by one
            //Action(Data, action);
        }

        public void CreateTickets(List<Dictionary<string, string>> TicketItemList)
        {
            //Access each ticket items
            for (int nIdx = 0; nIdx < TicketItemList.Count; nIdx++)
            {
                Dictionary<string, string> dItem = TicketItemList[nIdx];
                CreateOneTicket(dItem); // create ticket
            }
        }
    }
}
