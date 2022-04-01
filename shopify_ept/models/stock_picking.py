from odoo import models, fields, api, _


class stock_picking(models.Model):
    _inherit = "stock.picking"


    updated_in_shopify = fields.Boolean("Updated In Shopify", default=False)
    is_shopify_delivery_order = fields.Boolean("Shopify Delivery Order", store=True)
    shopify_instance_id = fields.Many2one("shopify.instance.ept", "Instance", store=True)
    is_cancelled_in_shopify =fields.Boolean("Is Cancelled In Shopify ?",default=False,copy=False,help="Use this field to identify shipped in Odoo but cancelled in Shopify")
    is_manually_action_shopify_fulfillment =fields.Boolean("Is Manually Action Required ?",default=False,copy=False,help="Those orders which we may fail update fulfillment status, we force set True and use will manually take necessary actions")
    shopify_fulfillment_id = fields.Char(string='Shopify Fulfillment Id')
