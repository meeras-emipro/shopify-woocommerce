import json
from odoo import models, fields, api, _
from datetime import datetime
from dateutil import parser
import pytz
from .. import shopify

utc = pytz.utc
import time
import logging
from odoo.exceptions import UserError
from ..shopify.pyactiveresource.connection import ClientError
from odoo.addons.shopify_ept.shopify.pyactiveresource.util import xml_to_dict

_logger = logging.getLogger("shopify_order_process===(Emipro):")


class SaleOrder(models.Model):
    _inherit = "sale.order"

    @api.depends('risk_ids')
    def _check_risk_order(self):
        for order in self:
            flag = False
            for risk in order.risk_ids:
                if risk.recommendation != 'accept':
                    flag = True
                    break
            order.is_risky_order = flag

    def _get_shopify_order_status(self):
        """
        Set updated_in_shopify of order from the pickings.
        @author: Maulik Barad on Date 06-05-2020.
        """
        for order in self:
            if order.shopify_instance_id:
                pickings = order.picking_ids.filtered(lambda x: x.state != "cancel")
                if pickings:
                    outgoing_picking = pickings.filtered(
                        lambda x: x.location_dest_id.usage == "customer")
                    if all(outgoing_picking.mapped("updated_in_shopify")):
                        order.updated_in_shopify = True
                        continue
                if order.state != 'draft' and order.moves_count > 0:
                    move_ids = self.env["stock.move"].search([("picking_id", "=", False),
                                                              ("sale_line_id", "in", order.order_line.ids)])
                    state = set(move_ids.mapped('state'))
                    if len(set(state)) == 1 and 'done' in set(state):
                        order.updated_in_shopify = True
                        continue
                order.updated_in_shopify = False
                continue
            order.updated_in_shopify = True

    def _search_shopify_order_ids(self, operator, value):
        query = """
                    select so.id from stock_picking sp
                    inner join sale_order so on so.procurement_group_id=sp.group_id                   
                    inner join stock_location on stock_location.id=sp.location_dest_id and stock_location.usage='customer'
                    where sp.updated_in_shopify %s true and sp.state != 'cancel'
                """ % (operator)
        if operator == '=':
            query += """union all
                    select so.id from sale_order as so
                    inner join sale_order_line as sl on sl.order_id = so.id
                    inner join stock_move as sm on sm.sale_line_id = sl.id
                    where sm.picking_id is NULL and sm.state = 'done' and so.shopify_instance_id notnull"""
        self._cr.execute(query)
        results = self._cr.fetchall()
        order_ids = []
        for result_tuple in results:
            order_ids.append(result_tuple[0])
        order_ids = list(set(order_ids))
        return [('id', 'in', order_ids)]

    shopify_order_id = fields.Char("Shopify Order Ref", copy=False)
    shopify_order_number = fields.Char("Shopify Order Number", copy=False)
    shopify_instance_id = fields.Many2one("shopify.instance.ept", "Instance", copy=False)
    shopify_order_status = fields.Char("Shopify Order Status", copy=False,
                                       help="Shopify order status when order imported in odoo at the moment order status in Shopify.")
    auto_workflow_process_id = fields.Many2one("sale.workflow.process.ept", "Auto Workflow", copy=False)
    shopify_payment_gateway_id = fields.Many2one('shopify.payment.gateway.ept',
                                                 string="Payment Gateway", copy=False)
    risk_ids = fields.One2many("shopify.order.risk", 'odoo_order_id', "Risks", copy=False)
    shopify_location_id = fields.Many2one("shopify.location.ept", "Shopify Location", copy=False)
    checkout_id = fields.Char("Checkout Id", copy=False)
    is_risky_order = fields.Boolean("Risky Order ?", compute=_check_risk_order, store=True)
    updated_in_shopify = fields.Boolean("Updated In Shopify ?", compute=_get_shopify_order_status,
                                        search='_search_shopify_order_ids')
    closed_at_ept = fields.Datetime("Closed At", copy=False)
    canceled_in_shopify = fields.Boolean("Canceled In Shopify", default=False, copy=False)
    is_service_tracking_updated = fields.Boolean("Service Tracking Updated", default=False, copy=False)
    is_pos_order = fields.Boolean("POS Order ?", copy=False)
    is_shopify_multi_payment = fields.Boolean("Multi Payments?", default=False, copy=False,
                                              help="It is used to identify that order has multi-payment gateway or not")
    shopify_payment_ids = fields.One2many('shopify.order.payment.ept', 'order_id',
                                          string="Payment Lines")

    _sql_constraints = [('unique_shopify_order',
                         'unique(shopify_instance_id,shopify_order_id,shopify_order_number)',
                         "Shopify order must be Unique.")]

    def update_warehouse_shopify_order(self, shopify_location, warehouse_id, pos_order):
        return {'shopify_location_id': shopify_location and shopify_location.id or False,
                "warehouse_id": warehouse_id,
                'is_pos_order': pos_order}

    def import_shopify_orders(self, order_data_queue_line, log_book_id):
        """This method used to create a sale orders in Odoo.
            @param : self, order_data_queue_line
            @return:
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 11/11/2019.
            Task Id : 157350
        """
        comman_log_line_obj = self.env["common.log.lines.ept"]
        res_partner_obj = self.env["res.partner"]
        shopify_location = shopify_location_obj = self.env["shopify.location.ept"]
        instance = order_data_queue_line.shopify_instance_id
        order_data = order_data_queue_line.order_data
        order_response = json.loads(order_data)
        order_response = order_response.get('order')
        model = "sale.order"
        model_id = comman_log_line_obj.get_model_id(model)
        instance.connect_in_shopify()
        shopify_financial_status = order_response.get("financial_status")
        _logger.info('Start process of shopify order(%s) and order id is(%s) '
                     % (order_response.get('order_number'), order_response.get('id')))
        order_id = self.search([('shopify_order_id', '=', order_response.get('id')),
                                ('shopify_instance_id', '=', instance.id),
                                ('shopify_order_number', '=', order_response.get('order_number'))], limit=1)
        if not order_id:
            order_id = self.search([
                ('shopify_instance_id', '=', instance.id),
                ('client_order_ref', '=', order_response.get('name'))], limit=1)

        if order_id:
            order_data_queue_line.write({'state': 'done', 'processed_at': datetime.now(),
                                         'sale_order_id': order_id.id})
            _logger.info('Done the Process of order Because Shopify Order(%s) is exist in Odoo and '
                         'Odoo order is(%s)' % (order_response.get('order_number'), order_id.name))
            self._cr.commit()
            return order_id

        if order_response.get("source_name", "") == "pos":
            if order_response.get("customer"):
                partner = res_partner_obj.create_shopify_pos_customer(order_response, log_book_id, instance)
            else:
                partner = instance.shopify_default_pos_customer_id
        else:
            # add by bhavesh jadav 05/12/2019
            parent_id = False
            ##Added below condition as per task 169296 and it will create a mismatch log if any address not found in order response.
            if not any([order_response.get('customer', False), order_response.get('shipping_address', False),
                        order_response.get('billing_address', False)]):
                message = "Customer Not Available In %s Order" % (order_response.get('order_number'))
                if order_response.get("source_name", "") == "pos":
                    message = "Default POS Customer is not set.\nPlease set Default POS Customer in Configuration."
                comman_log_line_obj.shopify_create_order_log_line(message, model_id,
                                                                  order_data_queue_line, log_book_id)
                order_data_queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
                _logger.info('Customer not available in Shopify Order(%s)' % (
                    order_response.get('order_number')))
                return False
            partner = order_response.get('customer', {}) and res_partner_obj.create_or_update_customer(
                vals=order_response,
                log_book_id=log_book_id,
                is_company=True,
                parent_id=False,
                type=False,
                instance=instance,
                order_data_queue_line=order_data_queue_line) or False
        if partner and partner.parent_id:
            parent_id = partner.parent_id
        if partner and not partner.parent_id:
            parent_id = partner
        # add by bhavesh jadav 05/12/2019
        shipping_address = order_response.get('shipping_address',
                                              False) and res_partner_obj.create_or_update_customer(
            vals=order_response.get('shipping_address'), log_book_id=log_book_id,
            is_company=False,
            parent_id=parent_id.id if parent_id else False, type='delivery', instance=instance,
            email=order_response.get('customer', {}).get('email')) or partner
        # add by bhavesh jadav 05/12/2019
        invoice_address = order_response.get('billing_address',
                                             False) and res_partner_obj.create_or_update_customer(
            vals=order_response.get('billing_address'), log_book_id=log_book_id,
            is_company=False,
            parent_id=parent_id.id if parent_id else False, type='invoice', instance=instance,
            email=order_response.get('customer', {}).get('email')) or partner
        # Added below condition as per task 169296.
        if not partner and invoice_address and shipping_address:
            partner = invoice_address
        if not partner and not shipping_address and invoice_address:
            partner = invoice_address
            shipping_address = invoice_address
        if not partner and not invoice_address and shipping_address:
            partner = shipping_address
            invoice_address = shipping_address

        lines = order_response.get('line_items')
        if self.check_mismatch_details(lines, instance, order_response.get('order_number'),
                                       order_data_queue_line, log_book_id):
            _logger.info('Mis-mismatch details found in this Shopify Order(%s) and id (%s)' % (
                order_response.get('order_number'), order_response.get('id')))
            order_data_queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
            return False
        shopify_location_id = order_response.get('location_id') or False
        if shopify_location_id:
            shopify_location = shopify_location_obj.search(
                [('shopify_location_id', '=', shopify_location_id),
                 ('instance_id', '=', instance.id)],
                limit=1)
        order_id = self.shopify_create_order(instance, partner, shipping_address, invoice_address,
                                             order_data_queue_line, order_response, log_book_id)
        if not order_id:
            _logger.info('Configuration missing in Odoo while importing Shopify Order(%s) '
                         'and id (%s)' % (
                             order_response.get('order_number'), order_response.get('id')))
            order_data_queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
            return False
        # add by Vrajesh Parekh 31/03/2020 for define POS order
        pos_order = True if order_response.get("source_name", "") == "pos" else False
        if shopify_location and shopify_location.warehouse_for_order:
            warehouse_id = shopify_location.warehouse_for_order.id
        else:
            warehouse_id = instance.shopify_warehouse_id.id
        order_id.write(order_id.update_warehouse_shopify_order(shopify_location, warehouse_id, pos_order))

        risk_result = shopify.OrderRisk().find(order_id=order_response.get('id'))
        if risk_result:
            self.env["shopify.order.risk"].shopify_create_risk_in_order(risk_result, order_id)
        _logger.info('Creating order line for Odoo order(%s) and Shopify order is (%s)' % (
            order_id.name, order_response.get('order_number')))
        total_discount = order_response.get('total_discounts', 0.0)
        for line in lines:
            shopify_product = self.search_shopify_product_for_order_line(line, instance)

            if not shopify_product:
                if line.get('fulfillment_service') == 'gift_card':
                    product = instance.gift_card_product_id
                if line.get('name') == 'Tip':
                    product = instance.tip_product_id
            if shopify_product:
                product = shopify_product.product_id

            order_line = self.shopify_create_sale_order_line(line, product, line.get('quantity'),
                                                             product.name, order_id,
                                                             line.get('price'), order_response)
            # add by Bhavesh Jadav 04/12/2019 for create separate  discount line fpr apply tax in discount line
            # add by Harsh Parekh  08/09/2020 for create seprate method for get discount allocation detail and call it below.Task Id : 166698
            if float(total_discount) > 0.0:
                self.create_discount_allocation_details_line(line, order_id, order_response, instance, product,
                                                             order_line)
        _logger.info('Created order line for Odoo order(%s) and Shopify order is (%s)' % (
            order_id.name, order_response.get('order_number')))

        # shipping line method
        self.create_or_update_shipping_lines_ept(instance, order_id, order_response)

        # create stock move for done qty
        if order_response.get('fulfillment_status') == 'partial':
            self.check_and_create_stock_move_based_on_mrp(order_id)

        # Add by Vrajesh Parekh -Dt: 31/03/2020 for pos order automatically delivered when import
        # sale order in shopify store.
        # order = order_id.filtered(lambda s: s.is_pos_order or s.shopify_order_status == "fulfilled")
        # if order:
        #     order.fulfilled_shopify_order()
        message = ""
        customer_loc = self.env['stock.location'].search([('usage', '=', 'customer')], limit=1)
        if order_response.get('fulfillment_status') == 'fulfilled':
            order_id.auto_workflow_process_id.shipped_order_workflow(order_id, customer_loc)
            # Below code add for create partially/fully refund
            message = self.create_shipped_order_refund(shopify_financial_status, order_response, order_id)
        else:
            self.env['sale.workflow.process.ept'].auto_workflow_process(ids=order_id.ids)

        _logger.info('Done auto workflow process for Odoo order(%s) and Shopify order is (%s)'
                     % (order_id.name, order_response.get('order_number')))
        if message and order_data_queue_line:
            model_id = comman_log_line_obj.get_model_id(self._name)
            comman_log_line_obj.shopify_create_order_log_line(message, model_id,
                                                              order_data_queue_line, log_book_id)
            order_data_queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
        else:
            if order_data_queue_line:
                order_data_queue_line.write({"state": "done", "processed_at": datetime.now(),
                                             "sale_order_id": order_id.id})
        order_data_queue_line.write({'state': 'done', 'processed_at': datetime.now(),
                                     'sale_order_id': order_id.id})
        _logger.info('Done the Odoo Order(%s) create process and Shopify Order (%s)' % (
            order_id.name, order_response.get('order_number')))
        return order_id

    def create_shipped_order_refund(self, shopify_financial_status, order_response, sale_order):
        """ This method is used to create partially or fully refund in shopify order.
            @param : shopify_financial_status, order_response, sale_order
            @return: message
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 30 December 2021 .
            Task_id: 181893
        """
        message = ""
        if shopify_financial_status == "refunded":
            total_refund = 0.0
            for refund in order_response.get('refunds'):
                # We take[0] because we got one transaction in one refund. If there are multiple refunds then
                # each transaction attaches with a refund.
                if refund.get('transactions') and refund.get('transactions')[0].get('kind') == \
                        'refund' and refund.get('transactions')[0].get('status') == 'success':
                    refunded_amount = refund.get('transactions')[0].get('amount')
                    total_refund += float(refunded_amount)
            refunded = sale_order.create_shopify_refund(order_response.get("refunds"), total_refund)
            if refunded[0] == 0:
                message = "- Refund can only be generated if it's related order " \
                          "invoice is found.\n- For order [%s], system could not find the " \
                          "related order invoice. " % (order_response.get('name'))
            elif refunded[0] == 2:
                message = "- Refund can only be generated if it's related order " \
                          "invoice is in 'Post' status.\n- For order [%s], system found " \
                          "related invoice but it is not in 'Post' status." % (
                              order_response.get('name'))
            elif refunded[0] == 3:
                message = "- Partial refund is received from Shopify for order [%s].\n " \
                          "- System do not process partial refunds.\n" \
                          "- Either create partial refund manually in Odoo or do full " \
                          "refund in Shopify." % (order_response.get('name'))
        elif shopify_financial_status == "partially_refunded" and order_response.get("refunds"):
            message = sale_order.create_shopify_partially_refund(order_response.get("refunds"),
                                                                 order_response.get('name'))

        return message

    def check_and_create_stock_move_based_on_mrp(self, order_id):
        """
        This method is used to created stock move based on parial order
        created stock move for done qty based on shopify fulfillable quantity
        so at the time of order confirmation picking is created for remaining qty
        @author: Nimesh Jethva @Emipro Technologies Pvt. Ltd on date 30/12/2020.
        Task Id : 169363
        """
        is_mrp_install = self.env['ir.module.module'].sudo().search([('name', '=', 'mrp'), ('state', '=', 'installed')])
        picking_obj = self.env['stock.picking']

        for line in order_id.order_line.filtered(
                lambda l: l.product_id.type != 'service' and l.shopify_fulfillable_quantity > 0):
            bom_lines = picking_obj.get_set_product(line.product_id)
            # if MRP is install and product is BOM product then move is created for all BOM
            if is_mrp_install and bom_lines:
                for bom_line in bom_lines:
                    self.shopify_create_and_done_stock_move_ept(line, order_id, bom_line)
            else:
                self.shopify_create_and_done_stock_move_ept(line, order_id)

    def shopify_create_and_done_stock_move_ept(self, line, order_id, bom_line=False):
        """prepare and create values for stock move"""
        customer_loc = self.env['stock.location'].search([('usage', '=', 'customer')], limit=1)

        product_qty = line.product_uom_qty - line.shopify_fulfillable_quantity
        product_id = bom_line[0].product_id.id if bom_line else line.product_id.id
        product_uom_qty = bom_line[1].get('qty', 0) * product_qty if bom_line else product_qty
        product_uom = bom_line[0].product_uom_id.id if bom_line else line.product_uom.id

        stock_move_values = {
            'name': ('Auto processed move : %s') % line.product_id.display_name,
            'company_id': order_id.company_id.id,
            'product_id': product_id or False,
            'product_uom_qty': product_uom_qty,
            'product_uom': product_uom or False,
            'location_id': order_id.warehouse_id.lot_stock_id.id,
            'location_dest_id': customer_loc.id,
            'state': 'confirmed',
            'sale_line_id': line.id
        }

        stock_move = self.env['stock.move'].create(stock_move_values)
        stock_move._action_assign()
        stock_move._set_quantity_done(product_uom_qty)
        stock_move._action_done()

        return True

    def create_discount_allocation_details_line(self, line, order_id, order_response, instance, product, order_line):
        """This method used to create discount line in order line.
            @param : discount_allocations
            @return:
            @author: Harsh Parekh @Emipro Technologies Pvt. Ltd on date 08/09/2020.
            Task Id : 166698
        """
        discount_amount = 0.0
        for discount_allocation in line.get('discount_allocations'):
            discount_amount += float(discount_allocation.get('amount'))
        if discount_amount > 0.0:
            _logger.info(
                'Creating discount line for Odoo order(%s) and Shopify order is (%s)'
                % (order_id.name, order_response.get('order_number')))
            self.shopify_create_sale_order_line({}, instance.discount_product_id, 1,
                                                product.name, order_id,
                                                float(discount_amount) * -1, order_response,
                                                previous_line=order_line, is_discount=True)
            _logger.info(
                'Created discount line for Odoo order(%s) and Shopify order is (%s)'
                % (order_id.name, order_response.get('order_number')))
        return True

    def create_or_update_shipping_lines_ept(self, instance, order_id, order_response):
        """This method used to create or update shipping line in order line.
            @param : instance, order_id , order_response
            @author: Harsh Parekh @Emipro Technologies Pvt. Ltd on date 10/09/2020.
            Task Id : 166698
        """
        for line in order_response.get('shipping_lines', []):
            carrier = self.env["delivery.carrier"].shopify_search_create_delivery_carrier(line, instance)
            order_id.write({'carrier_id': carrier.id if carrier else False})
            shipping_product = carrier.product_id
            self.shopify_create_sale_order_line(line, shipping_product, 1,
                                                shipping_product.name or line.get('title'),
                                                order_id, line.get('price'), order_response,
                                                is_shipping=True)
        _logger.info('Start auto workflow process for Odoo order(%s) and Shopify order is (%s)'
                     % (order_id.name, order_response.get('order_number')))
        return True

    def check_mismatch_details(self, lines, instance, order_number, order_data_queue_line,
                               log_book_id):
        """This method used to check the mismatch details in the order lines.
            @param : self, lines, instance, order_number, order_data_queue_line
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 11/11/2019.
            Task Id : 157350
        """
        odoo_product_obj = self.env['product.product']
        shopify_product_obj = self.env['shopify.product.product.ept']
        shopify_product_template_obj = self.env['shopify.product.template.ept']
        comman_log_line_obj = self.env["common.log.lines.ept"]
        model = "sale.order"
        model_id = comman_log_line_obj.get_model_id(model)
        mismatch = False
        for line in lines:
            if line.get('fulfillment_service') == 'gift_card' or line.get('name') == 'Tip':
                continue
            shopify_variant = False
            sku = line.get('sku') or False
            if line.get('variant_id', None):
                shopify_variant = shopify_product_obj.search(
                    [('variant_id', '=', line.get('variant_id')),
                     ('shopify_instance_id', '=', instance.id)])
            if not shopify_variant and sku:
                shopify_variant = shopify_product_obj.search(
                    [('default_code', '=', sku),
                     ('shopify_instance_id', '=', instance.id)])
            if shopify_variant:
                continue
            if not shopify_variant:
                line_variant_id = line.get('variant_id', False)
                line_product_id = line.get('product_id', False)
                if line_product_id and line_variant_id:
                    shopify_product_template_obj.shopify_sync_products(False, line_product_id,
                                                                       instance, log_book_id,
                                                                       order_data_queue_line)
                    if line.get('variant_id', None):
                        shopify_variant = shopify_product_obj.search(
                            [('variant_id', '=', line.get('variant_id')),
                             ('shopify_instance_id', '=', instance.id)])
                    if not shopify_variant and sku:
                        shopify_variant = shopify_product_obj.search(
                            [('default_code', '=', sku),
                             ('shopify_instance_id', '=', instance.id)])
                    if not shopify_variant:
                        message = "Product %s having %s not found for Order %s" % (
                            line.get('title'), line.get('title'), order_number)
                        comman_log_line_obj.shopify_create_order_log_line(message, model_id,
                                                                          order_data_queue_line,
                                                                          log_book_id)
                        mismatch = True
                        break
                else:
                    message = "Product Id Not Available In %s Order Line response" % (order_number)
                    comman_log_line_obj.shopify_create_order_log_line(message, model_id,
                                                                      order_data_queue_line,
                                                                      log_book_id)
                    order_data_queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
                    mismatch = True
                    break
        return mismatch

    def shopify_create_order(self, instance, partner, shipping_address, invoice_address,
                             order_data_queue_line, order_response, log_book_id):
        """This method used to create a sale order.
            @param : self, instance, partner, shipping_address, invoice_address,order_data_queue_line, order_response
            @return: order
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 12/11/2019.
            Task Id : 157350
        """
        payment_gateway_obj = self.env["shopify.payment.gateway.ept"]
        gateway = order_response.get('gateway') or "no_payment_gateway"
        payment_gateway, workflow = payment_gateway_obj.shopify_search_create_gateway_workflow(instance,
                                                                                               order_data_queue_line,
                                                                                               order_response,
                                                                                               log_book_id, gateway)

        if not all([payment_gateway, workflow]):
            return False

        order_vals = self.prepare_shopify_order_vals(instance, partner, shipping_address,
                                                     invoice_address, order_response,
                                                     payment_gateway,
                                                     workflow)
        if len(order_response.get('payment_gateway_names')) > 1:
            payment_vals = self.prepare_vals_shopify_multi_payment(instance, order_data_queue_line, order_response,
                                                                   log_book_id, payment_gateway, workflow)
            if not payment_vals:
                return False
            order_vals.update({'shopify_payment_ids': payment_vals, 'is_shopify_multi_payment': True})

        order = self.create(order_vals)
        return order

    def prepare_shopify_order_vals(self, instance, partner, shipping_address,
                                   invoice_address, order_response, payment_gateway,
                                   workflow):
        """
        This method used to Prepare a order vals.
        @param : self, instance, partner, shipping_address,invoice_address, order_response, payment_gateway,workflow
        @return: order_vals
        @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 13/11/2019.
        Task Id : 157350
        """
        date_order = self.convert_order_date(order_response)
        pricelist_id = self.shopify_set_pricelist(order_response=order_response, instance=instance)
        ordervals = {
            "company_id": instance.shopify_company_id.id if instance.shopify_company_id else False,
            "partner_id": partner.ids[0],
            "partner_invoice_id": invoice_address.ids[0],
            "partner_shipping_id": shipping_address.ids[0],
            "warehouse_id": instance.shopify_warehouse_id.id if instance.shopify_warehouse_id else False,
            "date_order": date_order,
            "state": "draft",
            "pricelist_id": pricelist_id.id if pricelist_id else False,
            "team_id": instance.shopify_section_id.id if instance.shopify_section_id else False,
        }
        ordervals = self.create_sales_order_vals_ept(ordervals)
        order_response_vals = self.prepare_order_vals_from_order_response(order_response, instance, workflow,
                                                                          payment_gateway)
        ordervals.update(order_response_vals)
        if not instance.is_use_default_sequence:
            if instance.shopify_order_prefix:
                name = "%s_%s" % (instance.shopify_order_prefix, order_response.get("name"))
            else:
                name = order_response.get("name")
            ordervals.update({"name": name})
        return ordervals

    def convert_order_date(self, order_response):
        """ This method is used to convert the order date in UTC and formate("%Y-%m-%d %H:%M:%S").
            @param order_response: Order response
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 30 December 2021.
            Task_id: 181893
        """
        if order_response.get("created_at", False):
            order_date = order_response.get("created_at", False)
            date_order = parser.parse(order_date).astimezone(utc).strftime("%Y-%m-%d %H:%M:%S")
        else:
            date_order = time.strftime("%Y-%m-%d %H:%M:%S")
            date_order = str(date_order)

        return date_order

    def prepare_order_vals_from_order_response(self, order_response, instance, workflow, payment_gateway):
        """ This method is used to prepare vals from the order response.
            @param order_response: Response of order.
            @param instance: Record of instance.
            @param workflow: Record of auto invoice workflow.
            @param payment_gateway: Record of payment gateway.
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 30 December 2021.
            Task_id: 181893
        """
        order_vals = {
            "checkout_id": order_response.get("checkout_id"),
            "note": order_response.get("note"),
            "shopify_order_id": order_response.get("id"),
            "shopify_order_number": order_response.get("order_number"),
            "shopify_payment_gateway_id": payment_gateway and payment_gateway.id or False,
            "shopify_instance_id": instance.id,
            "shopify_order_status": order_response.get("fulfillment_status") or "unfulfilled",
            "picking_policy": workflow.picking_policy or False,
            "auto_workflow_process_id": workflow and workflow.id,
            "client_order_ref": order_response.get("name")
        }
        return order_vals

    def shopify_set_pricelist(self, instance, order_response):
        """
        Author:Bhavesh Jadav 09/12/2019 for the for set price list based on the order response currency because of if
        order currency different then the erp currency so we need to set proper pricelist for that sale order
        otherwise set pricelist  based on instance configurations
        """
        currency_obj = self.env['res.currency']
        pricelist_obj = self.env['product.pricelist']
        order_currency = order_response.get('currency') or False
        if order_currency:
            currency = currency_obj.search([('name', '=', order_currency)])
            if not currency:
                currency = currency_obj.search(
                    [('name', '=', order_currency), ('active', '=', False)])
                if currency:
                    currency.write({'active': True})
                    pricelist = pricelist_obj.search([('currency_id', '=', currency.id)], limit=1)
                    if pricelist:
                        return pricelist
                    else:
                        pricelist_vals = {'name': currency.name,
                                          'currency_id': currency.id,
                                          'company_id': instance.shopify_company_id.id, }
                        pricelist = pricelist_obj.create(pricelist_vals)
                        return pricelist
                else:
                    pricelist = instance.shopify_pricelist_id if instance.shopify_pricelist_id else False
                    return pricelist
            else:
                pricelist = pricelist_obj.search([('currency_id', '=', currency.id)], limit=1)
                return pricelist
        else:
            pricelist = instance.shopify_pricelist_id if instance.shopify_pricelist_id else False
            return pricelist

    def search_shopify_product_for_order_line(self, line, instance):
        """This method used to search shopify product for order line.
            @param : self, line, instance
            @return: shopify_product
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 14/11/2019.
            Task Id : 157350
        """
        shopify_product_obj = self.env['shopify.product.product.ept']
        variant_id = line.get('variant_id')
        shopify_product = False
        if variant_id:
            shopify_product = shopify_product_obj.search(
                [('shopify_instance_id', '=', instance.id), ('variant_id', '=', variant_id),
                 ('exported_in_shopify', '=', True)], limit=1)
            if shopify_product:
                return shopify_product
        shopify_product = shopify_product_obj.search([('shopify_instance_id', '=', instance.id),
                                                      ('default_code', '=', line.get('sku')),
                                                      ('exported_in_shopify', '=', True)], limit=1)
        shopify_product and shopify_product.write({'variant_id': variant_id})
        if shopify_product:
            return shopify_product

    def shopify_create_sale_order_line(self, line, product, quantity,
                                       product_name, order_id,
                                       price, order_response, is_shipping=False,
                                       previous_line=False,
                                       is_discount=False):
        """This method used to create a sale order line.
            @param : self, line, product, quantity,product_name, order_id,price, is_shipping=False
            @return: order_line_id
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 14/11/2019.
            Task Id : 157350
        """
        sale_order_line_obj = self.env['sale.order.line']

        uom_id = product and product.uom_id and product.uom_id.id or False
        line_vals = {
            'product_id': product and product.ids[0] or False,
            'order_id': order_id.id,
            'company_id': order_id.company_id.id,
            'product_uom': uom_id,
            'name': product_name,
            'price_unit': price,
            'order_qty': quantity,
        }
        order_line_vals = sale_order_line_obj.create_sale_order_line_ept(line_vals)
        if order_id.shopify_instance_id.apply_tax_in_order == 'create_shopify_tax':
            taxes_included = order_response.get('taxes_included') or False
            tax_ids = []
            if line and line.get('tax_lines'):
                if line.get('taxable'):
                    # This is used for when the one product is taxable and another product is not
                    # taxable
                    tax_ids = self.shopify_get_tax_id_ept(order_id.shopify_instance_id,
                                                          line.get('tax_lines'),
                                                          taxes_included)
                if is_shipping:
                    # In the Shopify store there is configuration regarding tax is applicable on shipping or not, if applicable then this use.
                    tax_ids = self.shopify_get_tax_id_ept(order_id.shopify_instance_id,
                                                          line.get('tax_lines'),
                                                          taxes_included)
            elif not line and previous_line:
                # Before modification, connector set order taxes on discount line but as per connector design,
                # we are creating discount line base on sale order line so it should apply sale order line taxes
                # in discount line not order taxes. It creates a problem while the customer is using multi taxes in sale orders.
                # so set the previous line taxes on the discount line.
                tax_ids = [(6, 0, previous_line.tax_id.ids)]

            order_line_vals["tax_id"] = tax_ids
            # When the one order with two products one product with tax and another product
            # without tax and apply the discount on order that time not apply tax on discount
            # which is
            if is_discount and not previous_line.tax_id:
                order_line_vals["tax_id"] = []

        if is_discount:
            order_line_vals["name"] = 'Discount for ' + str(product_name)
            if order_id.shopify_instance_id.apply_tax_in_order == 'odoo_tax' and is_discount:  # add by bhavesh jadav 04/12/2019 for by pass odoo tax flow on discount line
                order_line_vals["tax_id"] = previous_line.tax_id

        order_line_vals.update({
            'shopify_line_id': line.get('id'),
            'is_delivery': is_shipping,
            'shopify_fulfillable_quantity': line.get('fulfillable_quantity'),
        })
        # if product is gift card then take name from shopify product
        if line.get('fulfillment_service') == 'gift_card':
            order_line_vals.update({'name': line.get('name')})
        order_line_id = sale_order_line_obj.create(order_line_vals)
        order_line_id.with_context(round=False)._compute_amount()
        return order_line_id

    @api.model
    def shopify_get_tax_id_ept(self, instance, tax_lines, tax_included):
        """This method used to search tax in Odoo.
            @param : self,instance,order_line,tax_included
            @return: tax_id
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 18/11/2019.
            Task Id : 157350
        """
        tax_id = []
        taxes = []
        company = instance.shopify_warehouse_id.company_id
        for tax in tax_lines:
            rate = float(tax.get('rate', 0.0))
            price = float(tax.get('price', 0.0))
            title = tax.get('title')
            rate = rate * 100
            if rate != 0.0 and price != 0.0:
                # add condition by Bhavesh Jadav 19/12/2019 for if the rate is same other details is same but name
                # its different then its apply wrong tax
                if tax_included:
                    name = '%s_(%s %s included)_%s' % (title, str(rate), '%', company.name)
                else:
                    name = '%s_(%s %s excluded)_%s' % (title, str(rate), '%', company.name)
                acctax_id = self.env['account.tax'].search(
                    [('price_include', '=', tax_included), ('type_tax_use', '=', 'sale'),
                     ('amount', '=', rate), ('name', '=', name),
                     ('company_id', '=', instance.shopify_warehouse_id.company_id.id)], limit=1)
                if not acctax_id:
                    acctax_id = self.shopify_create_account_tax(instance, rate, tax_included,
                                                                company, name)
                if acctax_id:
                    taxes.append(acctax_id.id)
        if taxes:
            tax_id = [(6, 0, taxes)]
        return tax_id

    @api.model
    def shopify_create_account_tax(self, instance, value, price_included, company, name):
        """This method used to create tax in Odoo when importing orders from Shopify to Odoo.
            @param : self, value, price_included, company, name
            @return: account_tax_id
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 18/11/2019.
            Task Id : 157350
        """
        account_tax_obj = self.env['account.tax']
        account_tax_id = account_tax_obj.create(
            {'name': name, 'amount': float(value), 'type_tax_use': 'sale',
             'price_include': price_included, 'company_id': company.id})
        account_tax_id.mapped('invoice_repartition_line_ids').write(
            {'account_id': instance.invoice_tax_account_id.id if
            instance.invoice_tax_account_id else False})
        account_tax_id.mapped('refund_repartition_line_ids').write(
            {'account_id': instance.credit_tax_account_id.id if instance.credit_tax_account_id
            else False})

        return account_tax_id

    def prepare_vals_shopify_multi_payment(self, instance, order_data_queue_line, order_response, log_book_id,
                                           payment_gateway, workflow):
        """ This method is use to prepare a values for the multi payment.
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 03/01/2022.
            Task_id: 181893 - Shopify fixes as per the new version.
        """
        payment_gateway_obj = self.env["shopify.payment.gateway.ept"]
        transactions = shopify.Transaction().find(order_id=order_response.get('id'))
        payment_list_vals = []
        for transaction in transactions:
            result = transaction.to_dict()
            if result.get('kind') in ['capture', 'sale'] and result.get('status') == 'success':
                payment_transaction_id = result.get('id')
                gateway = result.get('gateway')
                amount = result.get('amount')
                if order_response.get('gateway') == gateway:
                    payment_list = (0, 0, {'payment_gateway_id': payment_gateway.id, 'workflow_id': workflow.id,
                                           'amount': amount, 'payment_transaction_id': payment_transaction_id,
                                           'remaining_refund_amount': amount})
                    payment_list_vals.append(payment_list)
                    continue
                new_payment_gateway, new_workflow = payment_gateway_obj.shopify_search_create_gateway_workflow(instance,
                                                                                                               order_data_queue_line,
                                                                                                               order_response,
                                                                                                               log_book_id,
                                                                                                               gateway)
                if not all([new_payment_gateway, new_workflow]):
                    return False
                payment_list = (0, 0, {'payment_gateway_id': new_payment_gateway.id, 'workflow_id': new_workflow.id,
                                       'amount': amount, 'payment_transaction_id': payment_transaction_id,
                                       'remaining_refund_amount': amount})
                payment_list_vals.append(payment_list)
        return payment_list_vals

    @api.model
    def closed_at(self, instance):
        sales_orders = self.search([('warehouse_id', '=', instance.shopify_warehouse_id.id),
                                    ('shopify_order_id', '!=', False),
                                    ('shopify_instance_id', '=', instance.id),
                                    ('state', '=', 'done'), ('closed_at_ept', '=', False)],
                                   order='date_order')

        instance.connect_in_shopify()

        for sale_order in sales_orders:
            order = shopify.Order.find(sale_order.shopify_order_id)
            order.close()
            sale_order.write({'closed_at_ept': datetime.now()})
        return True

    def get_shopify_tracking_number_url(self, picking):
        if picking.carrier_tracking_url:
            return [picking.carrier_tracking_url]
        return []

    def get_shopify_carrier_code(self, picking):
        if picking.carrier_id:
            carrier_name = picking.carrier_id.shopify_tracking_company or picking.carrier_id.shopify_source or picking.carrier_id.name or ''
            return carrier_name
        return ''

    def update_order_status_in_shopify(self, instance):
        log_line_array = []
        comman_log_line_obj = self.env["common.log.lines.ept"]
        shopify_product_obj = self.env['shopify.product.product.ept']
        location_obj = self.env['stock.location']
        model = "sale.order"
        model_id = comman_log_line_obj.get_model_id(model)
        _logger.info(
            _("Update Order Status process start and you select '%s' Instance") % instance.name)
        stock_picking_obj = self.env['stock.picking']
        instance.connect_in_shopify()
        notify_customer = instance.notify_customer
        customer_locations = location_obj.search([('usage', '=', 'customer')])
        picking_ids = stock_picking_obj.search(
            [('shopify_instance_id', '=', instance.id), ('updated_in_shopify', '=', False),
             ('state', '=', 'done'),
             ('location_dest_id', 'in', customer_locations.ids), ('is_cancelled_in_shopify', '=', False)
                , ('is_manually_action_shopify_fulfillment', '=', False)],
            order='id desc').ids
        gift_card_product_id = instance.gift_card_product_id
        picking_count = 0
        for picking_id in picking_ids:
            picking = stock_picking_obj.browse(picking_id)
            new_fulfillment = False
            fulfillment_id = ''
            picking_count = picking_count + 1
            _logger.info(picking_count)
            if picking_count >= 20:
                self._cr.commit()
                picking_count = 0
                _logger.info("Commit")
            tracking_numbers = []
            line_items = {}
            carrier_name = self.get_shopify_carrier_code(picking)
            sale_order = picking.sale_id
            try:
                order = shopify.Order.find(sale_order.shopify_order_id)
                order_data = order.to_dict()
                _logger.info(order_data.get('fulfillment_status'))

                if order_data.get('fulfillment_status') == 'fulfilled':
                    _logger.info('Order is already fulfilled')
                    _logger.info(sale_order.name)
                    sale_order.picking_ids.filtered(lambda l: l.state == 'done').write({'updated_in_shopify': True})
                    continue
                elif order_data.get('cancelled_at') and order_data.get('cancel_reason'):
                    sale_order.picking_ids.filtered(lambda l: l.state == 'done').write(
                        {'is_cancelled_in_shopify': True})
                    continue
            except Exception as e:
                continue
            sale_order_line_ids = not sale_order.is_service_tracking_updated and \
                                  sale_order.order_line.filtered(lambda l: l.product_id.type == 'service'
                                                                           and l.shopify_line_id
                                                                           and l.product_id.id != gift_card_product_id.id
                                                                           and l.is_delivery == False).mapped(
                                      'shopify_line_id') or []
            _logger.info(("We are processing on picking ID : '%s'" % picking.id))

            _logger.info(("We are processing on sale order ID : '%s'" % sale_order.id))
            order_lines = sale_order.order_line
            if order_lines and order_lines.filtered(
                    lambda s: s.product_id.type != 'service' and s.shopify_line_id == False or ''):
                message = (_(
                    "Order status is not updated for order %s because shopify line id not found in this order." % picking.sale_id.name))
                _logger.info(message)
                log_line_array = shopify_product_obj.shopify_create_log(message, model_id, False,
                                                                        log_line_array)
                continue
            if picking.carrier_tracking_ref:
                manage_multi_tracking_number_in_shopify_delivery_order = False
            elif picking.mapped('package_ids').filtered(lambda l: l.tracking_no):
                manage_multi_tracking_number_in_shopify_delivery_order = True
            else:
                manage_multi_tracking_number_in_shopify_delivery_order = False

            if not manage_multi_tracking_number_in_shopify_delivery_order:
                not_kit_moves = picking.move_lines.filtered(
                    lambda x: x.sale_line_id.product_id.id == x.product_id.id and x.state == 'done')
                for move in not_kit_moves:
                    if move.sale_line_id and move.sale_line_id.shopify_line_id:
                        shopify_line_id = move.sale_line_id.shopify_line_id
                    line_items.update({shopify_line_id: {'tracking_no': [picking.carrier_tracking_ref],
                                                         'qty': move.product_qty,
                                                         }})
                kit_sale_line_ids = picking.move_lines.filtered(
                    lambda x: x.sale_line_id.product_id.id != x.product_id.id and x.state == 'done').mapped(
                    'sale_line_id')
                for kit_sale_line in kit_sale_line_ids:
                    if kit_sale_line and kit_sale_line.shopify_line_id:
                        shopify_line_id = kit_sale_line.shopify_line_id
                    line_items.update({shopify_line_id: {'tracking_no': [picking.carrier_tracking_ref],
                                                         'qty': kit_sale_line.product_uom_qty,
                                                         }})
            else:
                not_kit_moves = picking.move_lines.filtered(
                    lambda x: x.sale_line_id.product_id.id == x.product_id.id and x.state == 'done')
                for move in not_kit_moves:
                    tracking_number_wise_details = dict()
                    for move_line in move.move_line_ids:
                        qty = tracking_number_wise_details.get(
                            move_line.result_package_id.tracking_no or 'No_Tracking', 0)
                        tracking_number_wise_details.update(
                            {move_line.result_package_id.tracking_no or 'No_Tracking': (qty + move_line.qty_done)})
                    for tracking_no, qty in tracking_number_wise_details.items():
                        if tracking_no == 'No_Tracking':
                            tracking_no = ''
                        if move.sale_line_id and move.sale_line_id.shopify_line_id:
                            shopify_line_id = move.sale_line_id.shopify_line_id
                        line_items.update({shopify_line_id: {'tracking_no': [tracking_no],
                                                             'qty': qty,
                                                             }})
                kit_move_lines = picking.move_lines.filtered(
                    lambda x: x.sale_line_id.product_id.id != x.product_id.id and x.state == 'done')
                existing_sale_line_ids = []
                for move in kit_move_lines:
                    if move.sale_line_id.id in existing_sale_line_ids:
                        continue
                    existing_sale_line_ids.append(move.sale_line_id.id)
                    if move.sale_line_id and move.sale_line_id.shopify_line_id:
                        shopify_line_id = move.sale_line_id.shopify_line_id
                    tracking_no = move.move_line_ids.mapped('result_package_id').mapped('tracking_no') or []
                    tracking_no = tracking_no and tracking_no[0] or ''
                    line_items.update({shopify_line_id: {'tracking_no': [tracking_no],
                                                         'qty': move.product_qty,
                                                         }})
            update_lines = []

            for sale_line_id in line_items:
                tracking_numbers += line_items.get(sale_line_id).get('tracking_no')
                update_lines.append({'id': sale_line_id,
                                     'quantity': line_items.get(sale_line_id).get(
                                         'qty')})
            for line in sale_order_line_ids:
                quantity = sum(
                    sale_order.order_line.filtered(lambda l: l.shopify_line_id == line).mapped('product_uom_qty'))
                update_lines.append({'id': line, 'quantity': int(quantity)})
                sale_order.write({'is_service_tracking_updated': True})
            if not update_lines:
                message = "No lines found for update order status for %s" % (picking.name)
                _logger.info(message)
                log_line_array = shopify_product_obj.shopify_create_log(message, model_id, False,
                                                                        log_line_array)
                continue

            shopify_location_id = sale_order.shopify_location_id or False
            if not shopify_location_id:
                shopify_location_id = self.env['shopify.location.ept'].search([
                    ("warehouse_for_order", "=", sale_order.warehouse_id.id),
                    ('instance_id', '=', instance.id)])
                if shopify_location_id:
                    sale_order.shopify_location_id = shopify_location_id
                else:
                    shopify_location_id = self.env['shopify.location.ept'].search([
                        ('is_primary_location', '=', True),
                        ('instance_id', '=', instance.id)])
                if not shopify_location_id:
                    message = "Primary Location not found for instance %s while Update order status" % (
                        instance.name)
                    _logger.info(message)
                    log_line_array = shopify_product_obj.shopify_create_log(message, model_id,
                                                                            False,
                                                                            log_line_array)
                    continue
            try:
                new_fulfillment = shopify.Fulfillment(
                    {'order_id': sale_order.shopify_order_id,
                     'location_id': shopify_location_id.shopify_location_id,
                     'tracking_numbers': list(set(tracking_numbers)),
                     'tracking_urls': self.get_shopify_tracking_number_url(picking),
                     'tracking_company': carrier_name, 'line_items': update_lines,
                     'notify_customer': notify_customer})
            except Exception as e:
                if hasattr(e, 'response') and e.response.code == 429 and e.response.msg == "Too Many Requests":
                    time.sleep(int(float(e.response.headers.get('Retry-After', 5))))
                    new_fulfillment = shopify.Fulfillment(
                        {'order_id': sale_order.shopify_order_id,
                         'location_id': shopify_location_id.shopify_location_id,
                         'tracking_numbers': list(set(tracking_numbers)),
                         'tracking_urls': self.get_shopify_tracking_number_url(picking),
                         'tracking_company': carrier_name, 'line_items': update_lines,
                         'notify_customer': notify_customer})
            try:
                fulfillment_result = new_fulfillment.save()
            except ClientError as e:
                if hasattr(e, 'response') and e.response.code == 429 and e.response.msg == "Too " \
                                                                                           "Many Requests":
                    time.sleep(int(float(e.response.headers.get('Retry-After', 5))))
                    fulfillment_result = new_fulfillment.save()
            except Exception as e:
                message = "Order(%s) status not updated due to some issue in fulfillment " \
                          "request/response: %s" % (
                              sale_order.name, str(e))
                _logger.info(message)
                log_line_array = shopify_product_obj.shopify_create_log(message, model_id,
                                                                        False, log_line_array)
            if not fulfillment_result:
                if order_data.get('fulfillment_status') == 'partial':
                    picking.write({'updated_in_shopify': True})
                else:
                    picking.write({'is_manually_action_shopify_fulfillment': True})
                message = "Order(%s) status not updated due to some issue in fulfillment request/response:" % (
                    sale_order.name)
                _logger.info(message)
                log_line_array = shopify_product_obj.shopify_create_log(message, model_id,
                                                                        False, log_line_array)
                continue

            if new_fulfillment:
                shopify_fullment_result = xml_to_dict(new_fulfillment.to_xml())
                if shopify_fullment_result:
                    fulfillment_id = shopify_fullment_result.get('fulfillment').get('id') or ''
            picking.write({'updated_in_shopify': True, 'shopify_fulfillment_id': fulfillment_id})

        if len(log_line_array) > 0:
            shopify_product_obj.create_log(log_line_array, "export", instance)

        self.closed_at(instance)
        return True

    @api.model
    def process_shopify_order_via_webhook(self, order_data, instance, update_order=False):
        """
        Creates order data queue and process it.
        This method is for order imported via create and update webhook.
        @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 10-Jan-2020..
        @param order_data: Dictionary of order's data.
        @param instance: Instance of Shopify.
        """
        process_immediately = False if update_order else True

        order_data_queue = self.env[
            "shopify.order.data.queue.line.ept"].create_order_data_queue_line([order_data],
                                                                              instance,
                                                                              created_by='webhook',
                                                                              process_immediately=process_immediately)
        self._cr.commit()
        order_data_queue = self.env['shopify.order.data.queue.ept'].browse(order_data_queue)
        if order_data_queue:
            if not update_order:
                order_data_queue.order_data_queue_line_ids.process_import_order_queue_data()
                _logger.info("Imported order {0} of {1} via Webhook Successfully".format(
                    order_data.get("id"), instance.name))
        _logger.info(
            "Processed order {0} of {1} via Webhook Successfully".format(order_data.get("id"),
                                                                         instance.name))
        return True

    @api.model
    def update_shopify_order(self, queue_line, log_book):
        """
        This method will update order as per its status got from Shopify.
        @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 13-Jan-2020..
        @param queue_line: Order Data Queue Line.
        @param log_book: Common Log Book.
        @return: Updated Sale order.
        """
        message = ""
        shopify_instance = queue_line.shopify_instance_id
        order_response = json.loads(queue_line.order_data)
        order_data = order_response.get('order')
        shopify_status = order_data.get("financial_status")
        order = self.search([("shopify_instance_id", "=", shopify_instance.id),
                             ("shopify_order_id", "=", order_data.get("id"))])
        # Below condition use for, In shopify store there is full refund.
        if order_data.get('cancel_reason'):
            cancelled = order.cancel_shopify_order()
            if not cancelled:
                message = "System can not cancel the order {0} as one of the picking is in the done state.".format(
                    order.name)
        if shopify_status == "refunded":
            if not message:
                total_refund = 0.0
                for refund in order_data.get('refunds'):
                    # We take[0] because we got one transaction in one refund. If there are multiple refunds then each transaction attaches with a refund.
                    if refund.get('transactions') and refund.get('transactions')[0].get('kind') == \
                            'refund' and refund.get('transactions')[0].get('status') == 'success':
                        refunded_amount = refund.get('transactions')[0].get('amount')
                        total_refund += float(refunded_amount)
                refunded = order.create_shopify_refund(order_data.get("refunds"), total_refund)
                if refunded[0] == 0:
                    message = "System can not generate a refund as the invoice is not found. "
                    "Please first create an invoice."
                elif refunded[0] == 2:
                    message = "System can not generate a refund as the invoice is not posted. "
                    "Please first post the invoice."
                elif refunded[0] == 3:
                    message = "Currently partial refund is created in Shopify. Either create credit"
                    " note manual or refund fully in shopify."
        elif shopify_status == "partially_refunded" and order_data.get("refunds"):
            message = order.create_shopify_partially_refund(order_data.get("refunds"), order_data.get('name'))
        # Below condition use for, In shopify store there is fulfilled order.
        elif order_data.get('fulfillment_status') == 'fulfilled':
            fulfilled = order.fulfilled_shopify_order()
            if isinstance(fulfilled, bool) and not fulfilled:
                message = "System can not complete the picking as there is not enough quantity."
            elif not fulfilled:
                message = "System can not complete the picking as {0}".format(fulfilled)

        if message:
            comman_log_line_obj = self.env["common.log.lines.ept"]
            model = "sale.order"
            model_id = comman_log_line_obj.get_model_id(model)
            comman_log_line_obj.shopify_create_order_log_line(message, model_id,
                                                              queue_line, log_book)
            queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
        else:
            queue_line.state = "done"
        return order

    def cancel_shopify_order(self):
        """
        Cancelled the sale order when it is cancelled in Shopify Store with full refund.
        @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 13-Jan-2020..
        """
        if "done" in self.picking_ids.mapped("state"):
            return False
        self.action_cancel()
        self.canceled_in_shopify = True
        return True

    def create_shopify_refund(self, refunds_data, total_refund):
        """
        Creates refund of shopify order, when order is refunded in Shopify.
        It will need invoice created and posted for creating credit note in Odoo, otherwise it will
        create log and generate activity as per configuration.
        @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 13-Jan-2020..
        @param refunds_data: Data of refunds.
        @return:[0] : When no invoice is created.
                [1] : When invoice is not posted.
                [2] : When partial refund was made in Shopify.
                [True]:When credit notes are created or partial refund is done.
        """
        if not self.invoice_ids:
            return [0]
        invoices = self.invoice_ids.filtered(lambda x: x.type == "out_invoice")
        refunds = self.invoice_ids.filtered(lambda x: x.type == "out_refund")
        for invoice in invoices:
            if not invoice.state == "posted":
                return [2]
        if self.amount_total == total_refund and not refunds:
            move_reversal = self.env["account.move.reversal"].create({"refund_method": "cancel",
                                                                      "reason": "Refunded from "
                                                                                "shopify"
                                                                      if len(refunds_data) > 1 else
                                                                      refunds_data[0].get("note")})
            move_reversal.with_context({"active_model": "account.move",
                                        "active_ids": invoices.ids}).reverse_moves()
            return [True]
        return [3]

    def fulfilled_shopify_order(self):
        """
        If order is not confirmed yet, confirms it first.
        Make the picking done, when order will be fulfilled in Shopify.
        This method is used for Update order webhook.
        @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 13-Jan-2020..
        """
        if self.state not in ["sale", "done", "cancel"]:
            self.action_confirm()
        return self.fulfilled_picking_for_shopify(self.picking_ids.filtered(lambda x:
                                                                            x.location_dest_id.usage
                                                                            == "customer"))

    def fulfilled_picking_for_shopify(self, pickings):
        """
        It will make the pickings done.
        This method is used for Update order webhook.
        @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 13-Jan-2020..
        """
        for picking in pickings.filtered(lambda x: x.state not in ['cancel', 'done']):
            if picking.state != "assigned":
                if picking.move_lines.move_orig_ids:
                    completed = self.fulfilled_picking_for_shopify(
                        picking.move_lines.move_orig_ids.picking_id)
                    if not completed:
                        return False
                picking.action_assign()
                # # Add by Vrajesh Dt.01/04/2020 automatically validate delivery when import POS
                # order in shopify
                if picking.sale_id and (
                        picking.sale_id.is_pos_order or picking.sale_id.shopify_order_status == "fulfilled"):
                    for move_id in picking.move_ids_without_package:
                        picking.move_line_ids.create({
                            'product_id': move_id.product_id.id,
                            'product_uom_id': move_id.product_id.uom_id.id,
                            'qty_done': move_id.product_uom_qty,
                            'location_id': move_id.location_id.id,
                            'picking_id': picking.id,
                            'location_dest_id': move_id.location_dest_id.id,
                        })
                    picking.with_context(auto_processed_orders_ept=True).action_done()
                    return True
                if picking.state != "assigned":
                    return False
            result = picking.button_validate()
            if isinstance(result, dict):
                if result.get("res_model", "") == "stock.immediate.transfer":
                    immediate_transfer = self.env["stock.immediate.transfer"].browse(
                        result.get("res_id"))
                    immediate_transfer.process()
                elif result.get("res_model", "") == "stock.backorder.confirmation":
                    backorder = self.env["stock.backorder.confirmation"].browse(
                        result.get("res_id"))
                    backorder._process()
            else:
                return result
        return True

    def create_shopify_partially_refund(self, refunds_data, order_name):
        """This method is used to check the required validation before create
            a partial refund and call child methods for a partial refund.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 21/07/2021.
            Task Id : 175458 - Manage partial refund via webhook in v13
        """
        account_move_obj = self.env['account.move']
        message = False
        if not self.invoice_ids:
            message = "- Partially refund can only be generated if it's related order " \
                      "invoice is found.\n- For order [%s], system could not find the " \
                      "related order invoice. " % order_name
            return message
        invoices = self.invoice_ids.filtered(lambda x: x.type == "out_invoice")
        for invoice in invoices:
            if not invoice.state == "posted":
                message = "- Partially refund can only be generated if it's related order " \
                          "invoice is in 'Post' status.\n- For order [%s], system found " \
                          "related invoice but it is not in 'Post' status." % order_name
                return message
        for refund_data_line in refunds_data:
            existing_refund = account_move_obj.search([("shopify_refund_id", "=", refund_data_line.get('id')),
                                                       ("shopify_instance_id", "=", self.shopify_instance_id.id)])
            if existing_refund:
                continue
            new_move = self.create_move_and_delete_not_necessary_line(refund_data_line, invoices)
            if refund_data_line.get('order_adjustments'):
                self.create_refund_adjustment_line(refund_data_line.get('order_adjustments'), new_move)
            new_move.with_context(check_move_validity=False)._recompute_dynamic_lines()
        return message

    def create_move_and_delete_not_necessary_line(self, refunds_data, invoices):
        """This method is used to create a reverse move of invoice and delete the invoice lines from the newly
            created move which product not refunded in Shopify.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 27/07/2021.
            Task Id : 175458 - Manage partial refund via webhook in v13
        """
        delete_move_lines = self.env['account.move.line']
        account_move_obj = self.env['account.move']
        shopify_line_ids = []
        shopify_line_ids_with_qty = {}
        for refund_line in refunds_data.get('refund_line_items'):
            shopify_line_ids.append(refund_line.get('line_item_id'))
            shopify_line_ids_with_qty.update({refund_line.get('line_item_id'): refund_line.get('quantity')})

        move_reversal = self.env["account.move.reversal"].with_context(
            {"active_model": "account.move", "active_ids": invoices.ids}).create(
            {"refund_method": "refund", "reason": "Partially Refunded from shopify" if len(refunds_data) > 1 else
            refunds_data.get("note")})

        action_move_data = move_reversal.reverse_moves()
        new_move = account_move_obj.browse(action_move_data.get('res_id'))
        new_move.write({'is_refund_in_shopify': True, 'shopify_refund_id': refunds_data.get('id')})
        for new_move_line in new_move.invoice_line_ids:
            shopify_line_id = new_move_line.sale_line_ids.shopify_line_id
            if shopify_line_id and int(
                    shopify_line_id) not in shopify_line_ids or new_move_line.product_id == self.shopify_instance_id.discount_product_id:
                delete_move_lines += new_move_line
            else:
                new_move_line.with_context(check_move_validity=False).quantity = shopify_line_ids_with_qty.get(int(
                    shopify_line_id))

        new_move.message_post(body=_("Credit note generated by Webhook as Order partially "
                                     "refunded in Shopify. This credit note has been created from "
                                     "<a href=# data-oe-model=sale.order data-oe-id=%d>%s</a>") % (self.id, self.name))
        self.message_post(body=_(
            "Partially credit note created <a href=# data-oe-model=account.move data-oe-id=%d>%s</a> via webhook") % (
                                   new_move.id, new_move.name))

        if delete_move_lines:
            delete_move_lines.with_context(check_move_validity=False).unlink()
            new_move.with_context(check_move_validity=False)._recompute_tax_lines()
        return new_move

    def create_refund_adjustment_line(self, order_adjustments, move_ids):
        """This method is used to create an invoice line in a new move to manage the adjustment refund.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 27/07/2021.
            Task Id : 175458 - Manage partial refund via webhook in v13
        """
        account_move_line_obj = self.env['account.move.line']
        adjustment_product = self.env.ref('shopify_ept.shopify_refund_adjustment_product', False)
        adjustments_amount = 0.0
        for order_adjustment in order_adjustments:
            adjustments_amount += float(order_adjustment.get('amount', 0.0))
        if abs(adjustments_amount) > 0:
            move_vals = {'product_id': adjustment_product.id, 'quantity': 1, 'price_unit': abs(adjustments_amount),
                         'move_id': move_ids.id, 'partner_id': move_ids.partner_id.id,
                         'name': adjustment_product.display_name}
            new_move_vals = account_move_line_obj.new(move_vals)
            new_move_vals._onchange_product_id()
            new_vals = account_move_line_obj._convert_to_write(
                {name: new_move_vals[name] for name in new_move_vals._cache})
            new_vals.update({'quantity': 1, 'price_unit': abs(adjustments_amount), 'tax_ids': []})
            account_move_line_obj.with_context(check_move_validity=False).create(new_vals)

    def _prepare_invoice(self):
        """This method used set a shopify instance in customer invoice.
            @param : self
            @return: inv_val
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 20/11/2019.
            Task Id : 157911
        """
        inv_val = super(SaleOrder, self)._prepare_invoice()
        if self.shopify_instance_id:
            inv_val.update({'shopify_instance_id': self.shopify_instance_id.id, 'is_shopify_multi_payment':
                self.is_shopify_multi_payment})
        return inv_val

    def cancel_in_shopify(self):
        """This method used to open a wizard to cancel order in Shopify.
            @param : self
            @return: action
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 20/11/2019.
            Task Id : 157911
        """
        view = self.env.ref('shopify_ept.view_shopify_cancel_order_wizard')
        context = dict(self._context)
        context.update({'active_model': 'sale.order', 'active_id': self.id, 'active_ids': self.ids})
        return {
            'name': _('Cancel Order In Shopify'),
            'type': 'ir.actions.act_window',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'shopify.cancel.refund.order.wizard',
            'views': [(view.id, 'form')],
            'view_id': view.id,
            'target': 'new',
            'context': context
        }

    def _get_invoiceable_lines(self, final=False):
        """Inherited base method to manage tax rounding in the invoice.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 14 May 2021.
            176146 - fix changes of v14 to v13
        """
        if self.shopify_instance_id:
            self.env.context = dict(self._context)
            self.env.context.update({'round': False})
        invoiceable_lines = super(SaleOrder, self)._get_invoiceable_lines(final)
        return invoiceable_lines

    def paid_invoice_ept(self, invoices):
        """
        This method auto paid invoice based on auto workflow method.
        @author: Dipesh Tanna
        @param invoices: Recordset of Invoice.
        Updated by twinkalc to reconcile the created payment
        Migration done by twinkalc August 2020
        Override the auto invoice workflow library method here to create separate payment records.
        """
        self.ensure_one()
        account_payment_obj = self.env['account.payment']
        for invoice in invoices:
            if invoice.amount_residual:
                if self.is_shopify_multi_payment:
                    for payment in self.shopify_payment_ids:
                        vals = invoice.prepare_payment_dict(payment.workflow_id)
                        vals.update({'amount': payment.amount})
                        payment_id = account_payment_obj.create(vals)
                        payment_id.post()
                else:
                    vals = invoice.prepare_payment_dict(self.auto_workflow_process_id)
                    payment_id = account_payment_obj.create(vals)
                    payment_id.post()
        return True


class SaleOrderLine(models.Model):
    _inherit = "sale.order.line"

    shopify_line_id = fields.Char("Shopify Line", copy=False)
    shopify_fulfillable_quantity = fields.Float('Fulfillable Quantity', copy=False)

    def unlink(self):
        """
        @author: Haresh Mori on date:17/06/2020
        """
        for record in self:
            if record.order_id.shopify_order_id:
                msg = _(
                    "You can not delete this line because this line is Shopify order line and we need Shopify line id while we are doing update order status")
                raise UserError(msg)
        return super(SaleOrderLine, self).unlink()


class ImportShopifyOrderStatus(models.Model):
    _name = "import.shopify.order.status"
    _description = 'Order Status'

    name = fields.Char("Name")
    status = fields.Char("Status")
