import json
import time

from datetime import datetime, timedelta
from odoo import models, fields
# from datetime import datetime
from .. import shopify
from ..shopify.pyactiveresource.connection import ClientError
from odoo.exceptions import Warning, UserError


class ShopifyPaymentGateway(models.Model):
    _name = 'shopify.payment.gateway.ept'
    _description = "Shopify Payment Gateway"

    name = fields.Char("Name", help="Payment method name")
    code = fields.Char("Code", help="Payment method code given by Shopify")
    shopify_instance_id = fields.Many2one("shopify.instance.ept", required=True, string="Instance")
    active = fields.Boolean(default=True)

    def import_payment_gateway(self, instance):
        """
        This method import payment gateway through Order API.
        @param instance: Shopify Instance
        @author: Hardik Dhankecha on Date 15-Dec-2020.
        """
        to_date = datetime.now()
        from_date = to_date - timedelta(7)

        try:
            results = shopify.Order().find(status="any", updated_at_min=from_date,
                                           updated_at_max=to_date, fields=['gateway'], limit=250)
        except ClientError as error:
            if hasattr(error, "response"):
                if error.response.code == 429 and error.response.msg == "Too Many Requests":
                    time.sleep(int(float(error.response.headers.get('Retry-After', 5))))
                    results = shopify.Order().find(status="any", updated_at_min=from_date,
                                                   updated_at_max=to_date, fields=['gateway'], limit=250)
                else:
                    message = str(error.code) + "\n" + json.loads(error.response.body.decode()).get("errors")
                    raise UserError(message)
        except Exception as error:
            raise UserError(error)

        for result in results:
            result = result.to_dict()
            gateway = result.get('gateway') or "no_payment_gateway"
            self.search_or_create_payment_gateway(instance, gateway)

        return True

    def search_or_create_payment_gateway(self, instance, gateway_name):
        """
        This method searches for payment gateway and create it, if not found.
        @param instance: Shopify Instance.
        @param gateway_name: Payment gateway name.
        @author: Hardik Dhankechaon Date 15-Dec-2020.
        """
        shopify_payment_gateway = self.search([('code', '=', gateway_name),
                                               ('shopify_instance_id', '=', instance.id)], limit=1)
        if not shopify_payment_gateway:
            shopify_payment_gateway = self.create({'name': gateway_name,
                                                   'code': gateway_name,
                                                   'shopify_instance_id': instance.id})
        return shopify_payment_gateway

    def shopify_search_create_gateway_workflow(self, instance, order_data_queue_line,
                                               order_response, log_book_id, gateway):
        """This method used to search or create a payment gateway and workflow in odoo when importing orders from Shopify to Odoo.
            @param : self, instance, order_data_queue_line,order_response
            @return: gateway, workflow
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 12/11/2019.
            Task Id : 157350
        """
        comman_log_line_obj = self.env["common.log.lines.ept"]
        model = "sale.order"
        model_id = comman_log_line_obj.get_model_id(model)
        auto_workflow_id = False
        # gateway = order_response.get('gateway') or "no_payment_gateway"
        shopify_payment_gateway = self.search(
            [('code', '=', gateway), ('shopify_instance_id', '=', instance.id)], limit=1)
        if not shopify_payment_gateway:
            shopify_payment_gateway = self.create(
                {'name': gateway, 'code': gateway, 'shopify_instance_id': instance.id})
        workflow_config = self.env['sale.auto.workflow.configuration.ept'].search(
            [('shopify_instance_id', '=', instance.id),
             ('payment_gateway_id', '=', shopify_payment_gateway.id),
             ('financial_status', '=', order_response.get('financial_status'))])
        if not workflow_config:
            message = "Workflow Configuration not found for this order %s and payment gateway is " \
                      "'%s' and financial status is '%s'. You can see the auto workflow " \
                      "configuration here: Shopify => Configuration => Financial Status" % (
                          order_response.get('name'), gateway,
                          order_response.get('financial_status'))
            comman_log_line_obj.shopify_create_order_log_line(message, model_id,
                                                              order_data_queue_line, log_book_id)
            order_data_queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
            return shopify_payment_gateway, auto_workflow_id
        auto_workflow_id = workflow_config and workflow_config.auto_workflow_id or False
        if auto_workflow_id and not auto_workflow_id.picking_policy:
            message = "Please set the proper auto workflow configuration and name is : %s. You " \
                      "can see the auto " \
                      "workflow configuration here: Shopify => Configuration => Sale Auto " \
                      "workflow" % (auto_workflow_id.name)
            comman_log_line_obj.shopify_create_order_log_line(message, model_id,
                                                              order_data_queue_line, log_book_id)
            order_data_queue_line.write({'state': 'failed', 'processed_at': datetime.now()})
            auto_workflow_id = False
        return shopify_payment_gateway, auto_workflow_id
