from odoo import api, fields, models


class IrCron(models.Model):
    _inherit = "ir.cron"

    @api.model
    def update_existing_cron_for_woocommerce_ept(self):
        config_param_obj = self.env['ir.config_parameter']
        config_param = config_param_obj.get_param('woo_cron_status',False)
        if not config_param:
            product_parent_cron = self.env.ref('woo_commerce_ept.ir_cron_parent_woo_product_process_child_cron',False)
            order_parent_cron = self.env.ref('woo_commerce_ept.process_woo_order_data_queue_parent_cron',False)
            customer_parent_cron = self.env.ref('woo_commerce_ept.ir_cron_parent_woo_customer_process_child_cron',False)
            coupon_parent_cron = self.env.ref('woo_commerce_ept.process_woo_coupon_data_queue_parent_cron',False)

            if(product_parent_cron and product_parent_cron.interval_number < 5):
                product_parent_cron.write({"interval_number":5, "name":"WooCommerce: Process Product Queue"})

            if (order_parent_cron and order_parent_cron.interval_number < 5):
                order_parent_cron.write({"interval_number": 5, "name":"WooCommerce: Process Order Queue"})

            if (customer_parent_cron and customer_parent_cron.interval_number < 5):
                customer_parent_cron.write({"interval_number": 5, "name":"WooCommerce: Process Customer Queue"})

            if (coupon_parent_cron and coupon_parent_cron.interval_number < 5):
                coupon_parent_cron.write({"interval_number": 5, "name":"WooCommerce: Process Coupon Queue"})
            config_param_obj.set_param("woo_cron_status",1)
        return True
