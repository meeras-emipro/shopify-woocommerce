from odoo import models,fields,api,_

class StockMove(models.Model):
    _inherit="stock.move"

    def _get_new_picking_values(self):
        """We need this method to set Shopify Instance in Stock Picking"""
        res = super(StockMove,self)._get_new_picking_values()
        order_id=self.sale_line_id.order_id
        if order_id.shopify_order_id != False:
            order_id and res.update({'shopify_instance_id': order_id.shopify_instance_id.id,'is_shopify_delivery_order':True})
        return res

    def _action_assign(self):
        # We inherited the base method here to set the instance values in picking while the picking type is dropship.
        res = super(StockMove, self)._action_assign()
        picking_ids = self.mapped('picking_id')
        for picking in picking_ids:
            if not picking.shopify_instance_id and picking.sale_id and picking.sale_id.shopify_instance_id:
                picking.write({'shopify_instance_id':picking.sale_id.shopify_instance_id.id})
        return res
