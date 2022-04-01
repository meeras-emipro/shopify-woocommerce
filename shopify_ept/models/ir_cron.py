from odoo import api, fields, models


class IrCron(models.Model):
    _inherit = "ir.cron"

    @api.model
    def update_existing_cron_for_shopify_ept(self):
        config_param_obj = self.env['ir.config_parameter']
        config_param = config_param_obj.get_param('shopify_cron_status',False)
        if not config_param:
            product_parent_cron = self.env.ref('shopify_ept.ir_cron_parent_to_process_product_queue_line',False)
            order_parent_cron = self.env.ref('shopify_ept.ir_cron_parent_to_process_order_queue_data',False)
            customer_parent_cron = self.env.ref('shopify_ept.ir_cron_parent_to_process_shopify_synced_customer_data',False)

            if(product_parent_cron and product_parent_cron.interval_number < 5):
                product_parent_cron.write({"interval_number":5, "name":"Shopify: Process Product Queue"})

            if (order_parent_cron and order_parent_cron.interval_number < 5):
                order_parent_cron.write({"interval_number": 5, "name":"Shopify: Process Order Queue"})

            if (customer_parent_cron and customer_parent_cron.interval_number < 5):
                customer_parent_cron.write({"interval_number": 5, "name":"Shopify: Process Customer Queue"})
                
            config_param_obj.set_param("shopify_cron_status",1)
        return True
