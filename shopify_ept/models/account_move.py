from odoo import models, fields, api, _


class AccountMove(models.Model):
    _inherit = "account.move"

    is_refund_in_shopify = fields.Boolean("Refund In Shopify", default=False)
    shopify_instance_id = fields.Many2one("shopify.instance.ept", "Instances")
    shopify_refund_id = fields.Char(help="Id of shopify refund.", copy=False)
    is_shopify_multi_payment = fields.Boolean("Multi Payments?", default=False, copy=False,
                                              help="It is used to identify that order has multi-payment gateway or not")

    def refund_in_shopify(self):
        """This method used to open a wizard to Refund order in Shopify.
            @param : self
            @return: action
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 20/11/2019.
            Task Id : 157911
        """
        view = self.env.ref('shopify_ept.view_shopify_refund_wizard')
        context = dict(self._context)
        context.update({'active_model': 'account.invoice', 'active_id': self.id, 'active_ids': self.ids})
        if self.reversed_entry_id.is_shopify_multi_payment:
            payment_gateway_ids = self.reversed_entry_id.invoice_line_ids.sale_line_ids.order_id.shopify_payment_ids
            payment_gateway_ids.write({'refund_amount': 0.0, 'is_want_to_refund': False})
            remaining_to_refund_payment_ids = payment_gateway_ids.filtered(
                lambda payment: payment.is_fully_refunded == False).payment_gateway_id.ids
            context.update({'display_refund_from': False, 'payment_gateway_ids': remaining_to_refund_payment_ids,
                            'default_payment_ids': [(6, 0, payment_gateway_ids.ids)]})
        return {
            'name': _('Refund order In Shopify'),
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'shopify.cancel.refund.order.wizard',
            'views': [(view.id, 'form')],
            'view_id': view.id,
            'target': 'new',
            'context': context
        }

    def open_shopify_multi_payment(self):
        """This method used to open a wizard to display details of multi payments.
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 03/01/2022.
            Task Id : 181893 - Shopify fixes as per the new version
        """
        context = dict(self._context)
        context.update({'active_model': 'account.invoice', 'active_id': self.id, 'active_ids': self.ids})
        view_id = self.env.ref('shopify_ept.shopify_multi_payment_gateway_tree_view_ept').id
        return {
            'name': _('Multi payments'),
            'type': 'ir.actions.act_window',
            'view_mode': 'tree',
            'res_model': 'shopify.order.payment.ept',
            'view_id': view_id,
            'views': [(view_id, 'tree')],
            'domain': [('id', 'in', self.line_ids.sale_line_ids.order_id.shopify_payment_ids.ids)],
            "target": "new",
            'context': context
        }
