using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RFQEventReceiver.Entities.OrderItems
{
    /// <summary>
    /// Simple Factory class for instantiating OrderItems.
    /// </summary>
    public static class OrderItemFactory
    {
        /// <summary>
        /// Create a specific type of Order Item based upon the ContractType.
        /// </summary>
        /// <param name="contractType"></param>
        /// <returns>An OrderItem instance based on the specified ContractType.</returns>
        public static OrderItemBase CreateOrderItem(string contractType)
        {
            if (contractType == ContractType.TENT.ToString())
            {
                return new TENTOrderItem();
            }
            else if (contractType == ContractType.SOE.ToString())
            {
                return new SOEOrderItem();
            }
            else if (contractType == ContractType.FES.ToString())
            {
                return new FESOrderItem();
            }
            else
            {
                return null;
            }
        }
    }
}
