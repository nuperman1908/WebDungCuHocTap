using Models.EF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models.DAO
{
   public class OrderDetailDao
    {
        DBShop db = null;
        public OrderDetailDao()
        {
            db = new DBShop();
        }
        public bool Insert(OrderDetail orderDetail)
        {
            try
            {
                db.OrderDetails.Add(orderDetail);
                db.SaveChanges();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
