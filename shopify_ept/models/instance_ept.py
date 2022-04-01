import logging
from odoo import models, fields, api, _
from odoo.exceptions import Warning, UserError
from calendar import monthrange
from .. import shopify

_logger = logging.getLogger("Shopify : ")
_secondsConverter = {
    'days': lambda interval: interval * 24 * 60 * 60,
    'hours': lambda interval: interval * 60 * 60,
    'weeks': lambda interval: interval * 7 * 24 * 60 * 60,
    'minutes': lambda interval: interval * 60,
}


class ShopifyInstanceEpt(models.Model):
    _name = "shopify.instance.ept"
    _description = 'Shopify Instance'

    @api.model
    def _get_default_warehouse(self):
        warehouse = self.env['stock.warehouse'].search([('company_id', '=', self.env.company.id)],
                                                       limit=1, order='id')
        return warehouse.id if warehouse else False

    @api.model
    def _default_stock_field(self):
        stock_field = self.env['ir.model.fields'].search(
            [('model_id.model', '=', 'product.product'), ('name', '=', 'virtual_available')],
            limit=1)
        return stock_field.id if stock_field else False

    @api.model
    def _default_order_status(self):
        """
        Return default status of shopify order, for importing the particular orders having this
        status.
        @author: Haresh Mori on Date 16-Dec-2019.
        """
        order_status = self.env.ref('shopify_ept.unshipped')
        return [(6, 0, [order_status.id])] if order_status else False

    @api.model
    def _default_discount_product(self):
        """
        Gives default discount product to set in imported shopify order.
        @author: Haresh Mori on Date 16-Dec-2019.
        """
        discount_product = self.env.ref('shopify_ept.shopify_discount_product', False)
        return discount_product

    def _count_all(self):
        for instance in self:
            # Below is used for product count.
            product_query = self.prepare_query_to_count_record('shopify_product_template_ept', instance)
            instance.product_count = self.query_to_product_count(product_query)
            instance.exported_product_count = self.query_to_product_count(False, instance, 'exported_in_shopify',
                                                                          'true')
            instance.ready_to_expor_product_count = self.query_to_product_count(False, instance,
                                                                                'exported_in_shopify', 'false')
            instance.published_product_count = self.query_to_product_count(False, instance, 'website_published', 'true')
            unpublished_product_query = product_query + " and website_published = 'false' and exported_in_shopify = 'true'"
            instance.unpublished_product_count = self.query_to_product_count(unpublished_product_query)
            # Below is used for sale order count.
            sale_query = self.prepare_query_to_count_record('sale_order', instance)
            instance.sale_order_count = self.query_to_sale_order_count(sale_query)
            instance.quotation_count = self.query_to_sale_order_count(False, instance, 'state', ('draft', 'sent'))
            order_query = sale_query + " and state not in ('draft', 'sent', 'cancel')"
            instance.order_count = self.query_to_sale_order_count(order_query)
            risky_order_query = sale_query + " and state IN ('draft') and is_risky_order = true"
            instance.risk_order_count = self.query_to_sale_order_count(risky_order_query)
            # Below is used for picking count.
            picking_query = self.prepare_query_to_count_record('stock_picking', instance)
            instance.picking_count = self.query_to_delivery_count(picking_query)
            instance.confirmed_picking_count = self.query_to_delivery_count(False, instance, 'state', 'confirmed')
            instance.assigned_picking_count = self.query_to_delivery_count(False, instance, 'state', 'assigned')
            instance.done_picking_count = self.query_to_delivery_count(False, instance, 'state', 'done')
            # Below is used for invoice count.
            invoice_query = self.prepare_query_to_count_record('account_move', instance)
            instance.invoice_count = self.query_to_invoice_count(invoice_query)
            open_invoice_query = invoice_query + " and state='posted' and type='out_invoice' and invoice_payment_state != 'paid'"
            instance.open_invoice_count = self.query_to_invoice_count(open_invoice_query)
            paid_invoice_query = invoice_query + " and state='posted' and type='out_invoice' and invoice_payment_state = 'paid'"
            instance.paid_invoice_count = self.query_to_invoice_count(paid_invoice_query)
            refund_invoice_query = invoice_query + " and type='out_refund'"
            instance.refund_invoice_count = self.query_to_invoice_count(refund_invoice_query)

    def prepare_query_to_count_record(self, table_name, instance):
        """ This method is used to prepare a query.
            :param table_name: Name of table
            :param instance: Record of instance.
            @return: query
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 9 December 2020 .
            Task_id: 168799 - Shopify Dashboard changes v13 & v12
        """
        query = """select count(*) from %s where 
            shopify_instance_id=%s""" % (table_name, instance.id)
        return query

    def query_to_product_count(self, query, instance='', query_field='', value=''):
        """ This method is used to count the product record using the sql query.
            :param query: Sql query
            :param instance: Record of instance.
            :param query_field: field name
            :param value: value of field.
            @return: Count of record.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 8 December 2020 .
            Task_id: 168799 - Shopify Dashboard changes v13 & v12
        """
        if not query:
            query = """select count(*) from shopify_product_template_ept where shopify_instance_id=%s and 
            %s=%s""" % (instance.id, query_field, value)
        self._cr.execute(query)
        records = self._cr.dictfetchall()
        return records[0].get('count')

    def query_to_sale_order_count(self, query, instance='', query_field='', value=''):
        """ This method is used to count the sale order record using the sql query.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 8 December 2020 .
            Task_id: 168799 - Shopify Dashboard changes v13 & v12
        """
        if not query:
            query = """select count(*) from sale_order where shopify_instance_id=%s and 
            %s IN %s""" % (instance.id, query_field, value)
        self._cr.execute(query)
        records = self._cr.dictfetchall()
        return records[0].get('count')

    def query_to_delivery_count(self, query, instance='', query_field='', value=''):
        """ This method is used to count the delivery record using the sql query.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 8 December 2020 .
            Task_id: 168799 - Shopify Dashboard changes v13 & v12
        """
        if not query:
            query = """select count(*) from stock_picking where shopify_instance_id=%s and 
            %s='%s'""" % (instance.id, query_field, value)
        self._cr.execute(query)
        records = self._cr.dictfetchall()
        return records[0].get('count')

    def query_to_invoice_count(self, query):
        """ This method is used to count the delivery record using the sql query.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 8 December 2020 .
            Task_id: 168799 - Shopify Dashboard changes v13 & v12
        """
        self._cr.execute(query)
        records = self._cr.dictfetchall()
        return records[0].get('count')

    @api.model
    def _default_refund_adjustment_product(self):
        """
        This method is used to set the gift card product in an instance.
        @author: Meera Sidapara on Date 05-Jan-2022.
        """
        refund_adjustment_product = self.env.ref('shopify_ept.shopify_refund_adjustment_product', False)
        return refund_adjustment_product

    def _default_tip_product(self):
        """
        This method is used to set the tip product in an instance.
        @author: Haresh Mori on Date 17-June-2021.
        """
        tip_product = self.env.ref('shopify_ept.shopify_tip_product', False)
        return tip_product

    @api.model
    def set_default_tip_product_in_existing_instance(self):
        """ It is used to set tip product which has already created an instance.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 17 June 2021.
            Task_id: 172890 - Tip order improt v13
        """
        instances = self.search([])
        tip_product = self.env.ref('shopify_ept.shopify_tip_product', False)
        if instances and tip_product:
            instances.write({'tip_product_id': tip_product.id})

    name = fields.Char(size=120, string='Name', required=True)
    shopify_company_id = fields.Many2one('res.company', string='Company', required=True,
                                         default=lambda self:
                                         self.env.company)
    shopify_warehouse_id = fields.Many2one('stock.warehouse', string='Warehouse',
                                           default=_get_default_warehouse)
    shopify_pricelist_id = fields.Many2one('product.pricelist', string='Pricelist')
    shopify_order_prefix = fields.Char(size=10, string='Order Prefix',
                                       help="Enter your order prefix")
    shopify_fiscal_position_id = fields.Many2one('account.fiscal.position',
                                                 string='Fiscal Position')
    shopify_stock_field = fields.Many2one('ir.model.fields', string='Stock Field',
                                          default=_default_stock_field)
    shopify_country_id = fields.Many2one("res.country", "Country")
    shopify_api_key = fields.Char("API Key", required=True)
    shopify_password = fields.Char("Password", required=True)
    shopify_shared_secret = fields.Char("Secret Key", required=True)
    shopify_host = fields.Char("Host", required=True)
    is_image_url = fields.Boolean("Is Image URL?",
                                  help="Check this if you use Images from URL\nKeep as it is if you use Product images")
    inventory_adjustment_id = fields.Many2one('stock.inventory', "Last Inventory")
    shopify_active_instance = fields.Boolean(string="Is active instance")
    shopify_last_date_customer_import = fields.Datetime(string="Last Date Of Customer Import",
                                                        help="it is used to store last import customer date")
    shopify_last_date_update_stock = fields.Datetime(string="Last Date of Stock Update",
                                                     help="it is used to store last update inventory stock date")
    shopify_last_date_product_import = fields.Datetime(string="Last Date Of Product Import",
                                                       help="it is used to store last import product date")
    # is_bom_type_product = fields.Boolean(string="Manage BOM/Kit type products?",help="Manage BOM/Kit type product stock")
    auto_import_product = fields.Boolean(string="Auto Create Product if not found?")
    shopify_sync_product_with = fields.Selection([('sku', 'Internal Reference(SKU)'),
                                                  ('barcode', 'Barcode'),
                                                  ('sku_or_barcode',
                                                   'Internal Reference or Barcode'),
                                                  ], string="Sync Product With", default='sku')
    shopify_pricelist_id = fields.Many2one('product.pricelist', string='Pricelist')
    update_category_in_odoo_product = fields.Boolean(string="Update Category In Odoo Product ?",
                                                     default=False)
    shopify_stock_field = fields.Many2one('ir.model.fields', string='Stock Field')
    # payment_term_id = fields.Many2one('account.payment.term', string='Payment Term')
    last_date_order_import = fields.Datetime(string="Last Date Of Order Import",
                                             help="Last date of sync orders from Shopify to Odoo")
    shopify_section_id = fields.Many2one('crm.team', 'Sales Team')
    # global_channel_id = fields.Many2one('global.channel.ept', string="Global Channel")
    is_use_default_sequence = fields.Boolean("Use Odoo Default Sequence?",
                                             help="If checked,Then use default sequence of odoo while create sale order.")
    import_shopify_order_status_ids = fields.Many2many('import.shopify.order.status',
                                                       'shopify_instance_order_status_rel',
                                                       'instance_id', 'status_id',
                                                       "Import Order Status",
                                                       default=_default_order_status,
                                                       help="Select order status in which you want to import the orders from Shopify to Odoo.")
    # Account field
    shopify_store_time_zone = fields.Char("Store Time Zone",
                                          help='This field used to import order process')
    default_active_instance = fields.Boolean(string="Set as default active instance",
                                             help="This used to set a default active instance. "
                                                  "Base on these fields set the default values in the operation wizard. Default insatnce is only one")
    discount_product_id = fields.Many2one("product.product", "Discount",
                                          domain=[('type', '=', 'service'), ],
                                          default=_default_discount_product,
                                          help="This is used for set discount product in a sale order lines")
    gift_card_product_id = fields.Many2one(
        'product.product',
        string="Gift Card", domain=[('type', '=', 'service')],
        default=lambda self: self.env.ref('shopify_ept.shopify_gift_card_product', raise_if_not_found=False))

    shopify_shipping_product_id = fields.Many2one("product.product", "Shipping Product",
                                                  domain=[('type', '=', 'service'), ],
                                                  help="This is used for set shipping product in a sale order lines",
                                                  default=lambda self: self.env.ref(
                                                      'shopify_ept.shopify_shipping_product', False))
    # Auto cron field
    shopify_auto_cron = fields.Boolean(string="Shopify Auto Cron",
                                       help="Configure Auto cron")
    apply_tax_in_order = fields.Selection(
        [("odoo_tax", "Odoo Default Tax Behaviour"), ("create_shopify_tax",
                                                      "Create new tax If Not Found")],
        copy=False, help=""" For Shopify Orders :- \n
                    1) Odoo Default Tax Behaviour - The Taxes will be set based on Odoo's
                                 default functional behaviour i.e. based on Odoo's Tax and Fiscal Position configurations. \n
                    2) Create New Tax If Not Found - System will search the tax data received 
                    from Shopify in Odoo, will create a new one if it fails in finding it.""")
    invoice_tax_account_id = fields.Many2one('account.account', string='Invoice Tax Account')
    credit_tax_account_id = fields.Many2one('account.account', string='Credit Tax Account')
    auto_closed_order = fields.Boolean("Auto Closed Order", default=False)
    notify_customer = fields.Boolean("Notify Customer about Update Order Status?",
                                     help="If checked,Notify the customer via email about Update Order Status")
    color = fields.Integer(string='Color Index')

    # fields for kanban view
    product_count = fields.Integer(compute='_count_all', string="Product")

    sale_order_count = fields.Integer(compute='_count_all', string="Sale Order Count")

    picking_count = fields.Integer(compute='_count_all', string="Picking")

    invoice_count = fields.Integer(compute='_count_all', string="Invoice")

    exported_product_count = fields.Integer(compute='_count_all', string="Exported Products")

    ready_to_expor_product_count = fields.Integer(compute='_count_all', string="Ready For Export")

    published_product_count = fields.Integer(compute='_count_all', string="Published Product")

    unpublished_product_count = fields.Integer(compute='_count_all', string="#UnPublished Product")

    quotation_count = fields.Integer(compute='_count_all', string="Quotation")

    order_count = fields.Integer(compute='_count_all', string="Sales Orders")

    risk_order_count = fields.Integer(compute='_count_all', string="Risky Orders")

    confirmed_picking_count = fields.Integer(compute='_count_all', string="Confirm Picking")

    assigned_picking_count = fields.Integer(compute='_count_all', string="Assigned Pickings")

    done_picking_count = fields.Integer(compute='_count_all', string="Done Picking")

    open_invoice_count = fields.Integer(compute='_count_all', string="Open Invoice")

    paid_invoice_count = fields.Integer(compute='_count_all', string="Paid Invoice")

    refund_invoice_count = fields.Integer(compute='_count_all', string="Refund Invoices")
    shopify_global_channel_id = fields.Many2one('global.channel.ept', string="Global Channel")

    shopify_user_ids = fields.Many2many('res.users', 'shopify_instance_ept_res_users_rel',
                                        'res_config_settings_id', 'res_users_id',
                                        string='Responsible User')  # add by bhavesh jadav 28/11/2019 for set  responsible user for schedule activity
    shopify_activity_type_id = fields.Many2one('mail.activity.type',
                                               string="Activity Type")  # add by bhavesh jadav 28/11/2019 for set activity for schedule activity
    shopify_date_deadline = fields.Integer('Deadline lead days',
                                           help="its add number of  days in schedule activity deadline date ")  # add by bhavesh jadav 28/11/2019 for add date_deadline lead days user for schedule activity
    is_shopify_create_schedule = fields.Boolean("Create Schedule Activity ? ", default=False,
                                                help="If checked, Then Schedule Activity create on order data queues"
                                                     " will any queue line failed.")  # add by bhavesh jadav 04/12/2019
    active = fields.Boolean("Active", default=True)
    sync_product_with_images = fields.Boolean("Sync Images?",
                                              help="Check if you want to import images along with "
                                                   "products",
                                              default=True)  # add by bhavesh jadav 17/12/2019 for images

    webhook_ids = fields.One2many("shopify.webhook.ept", "instance_id", "Webhooks")
    create_shopify_products_webhook = fields.Boolean("Manage Products via Webhooks",
                                                     help="True : It will create all product related webhooks.\nFalse : All product related webhooks will be deactivated.")

    create_shopify_customers_webhook = fields.Boolean("Manage Customers via Webhooks",
                                                      help="True : It will create all customer related webhooks.\nFalse : All customer related webhooks will be deactivated.")
    create_shopify_orders_webhook = fields.Boolean("Manage Orders via Webhooks",
                                                   help="True : It will create all order related webhooks.\nFalse : All "
                                                        "order related webhooks will be deactivated.")

    shopify_default_pos_customer_id = fields.Many2one("res.partner", "Default POS customer",
                                                      help="This customer will be set in POS order, when"
                                                           "customer is not found.")
    # Shopify Payout Report
    shopify_api_url = fields.Char(string="Payout API URL")
    transaction_line_ids = fields.One2many("shopify.payout.account.config.ept", "instance_id",
                                           string="Transaction Line")
    shopify_settlement_report_journal_id = fields.Many2one('account.journal',
                                                           string='Payout Report Journal')
    payout_last_import_date = fields.Date(string="Last Date of Payout Import")
    tip_product_id = fields.Many2one("product.product", "TIP", domain=[('type', '=', 'service')],
                                     default=_default_tip_product,
                                     help="This is used for set tip product in a sale order lines")
    refund_adjustment_product_id = fields.Many2one("product.product", "Refund Adjustment",
                                                   domain=[('type', '=', 'service')],
                                                   default=_default_refund_adjustment_product,
                                                   help="This is used for set refund adjustment in a credit note")

    @api.model
    def create(self, vals):
        """
        Inherited for creating generic POS customer.
        @author: Maulik Barad on date 25-Feb-2020.
        """
        customer_vals = {"name": "POS Customer(%s)" % vals.get("name"), "customer_rank": 1}
        customer = self.env["res.partner"].create(customer_vals)
        vals.update({"shopify_default_pos_customer_id": customer.id})
        return super(ShopifyInstanceEpt, self).create(vals)

    @api.onchange('default_active_instance')
    def _onchange_default_active_insatnce(self):
        default_active_instances = self.search([('default_active_instance', '=', True)])
        for default_active_instance in default_active_instances:
            if self and self.ids == default_active_instance.ids:
                default_active_instance.default_active_instance = True
            else:
                default_active_instance.default_active_instance = False

    def shopify_test_connection(self):
        """This method used to check the connection between Odoo and Shopify.
            @param : self
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 04/10/2019.
        """
        shop = self.shopify_host.split("//")
        if len(shop) == 2:
            shop_url = shop[0] + "//" + self.shopify_api_key + ":" + self.shopify_password + "@" + \
                       shop[1] + "/admin/api/2022-01"
        else:
            shop_url = "https://" + self.shopify_api_key + ":" + self.shopify_password + "@" + shop[
                0] + "/admin/api/2022-01"
        shopify.ShopifyResource.set_site(shop_url)
        try:
            shop_id = shopify.Shop.current()
        except Exception as e:
            raise Warning(e)
        shop_detail = shop_id.to_dict()
        self.write({"shopify_store_time_zone": shop_detail.get("iana_timezone")})
        title = _("Connection Test Succeeded!")
        message = _("Your connection is proper working!")
        self.env['bus.bus'].sendone(
            (self._cr.dbname, 'res.partner', self.env.user.partner_id.id),
            {
                'type': 'simple_notification', 'title': title, 'message': message, 'sticky': False,
                'warning': False
            })

    def connect_in_shopify(self):
        """This method used to connect with Odoo to Shopify.
            @param : self
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 07/10/2019.
        """
        instance = self
        shop = instance.shopify_host.split("//")
        if len(shop) == 2:
            shop_url = shop[
                           0] + "//" + instance.shopify_api_key + ":" + instance.shopify_password + "@" + \
                       shop[1] + "/admin/api/2022-01"
        else:
            shop_url = "https://" + instance.shopify_api_key + ":" + instance.shopify_password + "@" + \
                       shop[0] + "/admin/api/2022-01"

        shopify.ShopifyResource.set_site(shop_url)
        return True

    def shopify_action_archive_unarchive(self):
        """This method used to confirm the shopify instance.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 07/10/2019.
        """
        domain = [("shopify_instance_id", "=", self.id)]
        shopify_webhook_obj = self.env["shopify.webhook.ept"]
        shopify_location_obj = self.env["shopify.location.ept"]
        data_queue_mixin_obj = self.env['data.queue.mixin.ept']
        shopify_template_obj = self.env["shopify.product.template.ept"]
        sale_auto_workflow_configuration_obj = self.env["sale.auto.workflow.configuration.ept"]
        shopify_payment_gateway_obj = self.env["shopify.payment.gateway.ept"]
        if self.active:
            activate = {"active": False}
            domain_for_webhook_location = [("instance_id", "=", self.id)]
            self.write(activate)
            self.change_auto_cron_status(self)
            shopify_webhook_obj.search(domain_for_webhook_location).unlink()
            shopify_location_obj.search(domain_for_webhook_location).write(activate)
            data_queue_mixin_obj.delete_data_queue_ept(is_delete_queue=True)
        else:
            self.shopify_test_connection()
            activate = {"active": True}
            domain.append(("active", "=", False))
            self.write(activate)
            self.env['shopify.location.ept'].import_shopify_locations(self)

        shopify_template_obj.search(domain).write(activate)
        sale_auto_workflow_configuration_obj.search(domain).write(activate)
        shopify_payment_gateway_obj.search(domain).write(activate)
        return True

    def change_auto_cron_status(self, instance):
        """
        After connect or disconnect the shopify instance disable all the Scheduled Actions.
        :param instance:
        :return:
        @author: Angel Patel @Emipro Technologies Pvt. Ltd.
        Task Id : 157716
        """
        try:
            stock_cron_exist = self.env.ref(
                'shopify_ept.ir_cron_shopify_auto_export_inventory_instance_%d' % (instance.id))
        except:
            stock_cron_exist = False
        try:
            order_cron_exist = self.env.ref(
                'shopify_ept.ir_cron_shopify_auto_import_order_instance_%d' % (instance.id))
        except:
            order_cron_exist = False
        try:
            order_status_cron_exist = self.env.ref(
                'shopify_ept.ir_cron_shopify_auto_update_order_status_instance_%d' % (
                    instance.id))
        except:
            order_status_cron_exist = False

        if stock_cron_exist:
            stock_cron_exist.write({'active': False})
        if order_cron_exist:
            order_cron_exist.write({'active': False})
        if order_status_cron_exist:
            order_status_cron_exist.write({'active': False})

    def cron_configuration_action(self):
        """
        Open wizard from "Configure Scheduled Actions" button click
        @author: Angel Patel @Emipro Technologies Pvt. Ltd.
        Task Id : 157716
        :return:
        """
        action = self.env.ref('shopify_ept.action_wizard_shopify_cron_configuration_ept').read()[0]
        context = {
            'shopify_instance_id': self.id
        }
        action['context'] = context
        return action

    def action_redirect_to_ir_cron(self):
        """
        Redirect to ir.cron model with cron name like shopify
        @author: Angel Patel @Emipro Technologies Pvt. Ltd.
        Task Id : 157716
        :return:
        """
        action = self.env.ref('base.ir_cron_act').read()[0]
        action['domain'] = [('name', 'ilike', self.name)]
        return action

    def list_of_topic_for_webhook(self, event):
        """
        This method is prepare the list of all the event topic while the webhook create, and return that list
        :param event: having 'product' or 'customer' or 'order'
        :return: topic_list
        @author: Angel Patel on Date 17/01/2020.
        """
        if event == 'product':
            topic_list = ["products/create", "products/update", "products/delete"]
        if event == 'customer':
            topic_list = ["customers/create", "customers/update"]
        if event == 'order':
            topic_list = ["orders/create", "orders/updated", "orders/cancelled"]
        return topic_list

    def configure_shopify_product_webhook(self):
        """
        Creates or activates all product related webhooks, when it is True.
        Inactive all product related webhooks, when it is False.
        @author: Haresh Mori on Date 09-Jan-2020.
        :Modify by Angel Patel on date 17/01/2020, call list_of_topic_for_webhook method for get 'product' list events
        """
        topic_list = self.list_of_topic_for_webhook("product")
        self.configure_webhooks(topic_list)

    def configure_shopify_customer_webhook(self):
        """
        Creates or activates all customer related webhooks, when it is True.
        Inactive all customer related webhooks, when it is False.
        @author: Angel Patel on Date 10/01/2020.
        :Modify by Angel Patel on date 17/01/2020, call list_of_topic_for_webhook method for get 'customer' list events
        """
        topic_list = self.list_of_topic_for_webhook("customer")
        self.configure_webhooks(topic_list)

    def configure_shopify_order_webhook(self):
        """
        Creates or activates all order related webhooks, when it is True.
        Inactive all order related webhooks, when it is False.
        @author: Haresh Mori on Date 10/01/2020.
        :Modify by Angel Patel on date 17/01/2020, call list_of_topic_for_webhook method for get 'order' list events
        """
        topic_list = self.list_of_topic_for_webhook("order")
        self.configure_webhooks(topic_list)

    def configure_webhooks(self, topic_list):
        """
        Creates or activates all webhooks as per topic list, when it is True.
        Pauses all product related webhooks, when it is False.
        @author: Haresh Mori on Date 09/01/2020.
        """
        webhook_obj = self.env["shopify.webhook.ept"]

        resource = topic_list[0].split('/')[0]
        instance_id = self.id
        available_webhooks = webhook_obj.search(
            [("webhook_action", "in", topic_list), ("instance_id", "=", instance_id)])

        # self.refresh_webhooks(available_webhooks)

        if getattr(self, "create_shopify_%s_webhook" % (resource)):
            if available_webhooks:
                available_webhooks.write({'state': 'active'})
                _logger.info("{0} Webhooks are activated of instance '{1}'.".format(resource, self.name))
                topic_list = list(set(topic_list) - set(available_webhooks.mapped("webhook_action")))

            for topic in topic_list:
                webhook_obj.create({"webhook_name": self.name + "_" + topic.replace("/", "_"),
                                    "webhook_action": topic, "instance_id": instance_id})
                _logger.info("Webhook for '{0}' of instance '{1}' created.".format(topic, self.name))
        else:
            if available_webhooks:
                available_webhooks.write({'state': 'inactive'})
                _logger.info("{0} Webhooks are paused of instance '{1}'.".format(resource, self.name))

    def refresh_webhooks(self):
        """
        This method is used for delete record from the shopify.webhook.ept model record if webhook deleted from the shopify with some of the reasons.
        @author: Angel Patel@Emipro Technologies Pvt. Ltd on Date 15/01/2020.
        """
        self.connect_in_shopify()
        shopify_webhook = shopify.Webhook()
        responses = shopify_webhook.find()
        webhook_ids = []
        for webhook in responses:
            webhook_ids.append(str(webhook.id))
        _logger.info("Emipro-Webhook: Current webhook present in shopify is %s" % webhook_ids)
        webhook_obj = self.env['shopify.webhook.ept'].search(
            [('instance_id', '=', self.id), ('webhook_id', 'not in', webhook_ids)])
        _logger.info("Emipro-Webhook: Webhook not present in odoo is %s" % webhook_obj)
        if webhook_obj:
            for webhooks in webhook_obj:
                _logger.info("Emipro-Webhook: deleting the %s shopify.webhook.ept record" % webhooks.id)
                self._cr.execute("DELETE FROM shopify_webhook_ept WHERE id = %s", [webhooks.id], log_exceptions=False)
        _logger.info("Emipro-Webhook: refresh process done")
        return True

    def action_archive(self):
        action = self[0].with_context(active_ids=self.ids).action_shopify_active_archive_instance() if self else True
        return action

    def toggle_active(self):
        """ Method overridden for archiving the instance from the action menu.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 6 January 2021 .
            Task_id: 169828 - Changes of Active/Inactive instance in v13
        """
        action = self[0].with_context(active_ids=self.ids).action_shopify_active_archive_instance() if self else True
        return action

    def action_shopify_active_archive_instance(self):
        """ This method is used to open a wizard to display the information related to how many data will be
            archived/deleted while instance Active/Archive.
            @author: Haresh Mori @Emipro Technologies Pvt. Ltd on date 6 January 2021 .
            Task_id: 169828 - Changes of Active/Inactive instance in v13
        """
        view = self.env.ref('shopify_ept.view_active_archive_shopify_instance')
        return {
            'name': _('Instance Active/Archive Details'),
            'type': 'ir.actions.act_window',
            'res_model': 'shopify.queue.process.ept',
            'views': [(view.id, 'form')],
            'target': 'new',
            'context': self._context,
        }

    def get_shopify_cron_execution_time(self, cron_name):
        """
        This method is used to get the interval time of the cron.
        @param cron_name: External ID of the Cron.
        @return: Interval time in seconds.
        @author: Maulik Barad on Date 25-Nov-2020.
        """
        process_queue_cron = self.env.ref(cron_name, False)
        if not process_queue_cron:
            raise UserError(_("Please upgrade the module. \n Maybe the job has been deleted, it will be recreated at "
                              "the time of module upgrade."))
        interval = process_queue_cron.interval_number
        interval_type = process_queue_cron.interval_type
        if interval_type == "months":
            days = 0
            current_year = fields.Date.today().year
            current_month = fields.Date.today().month
            for i in range(0, interval):
                month = current_month + i

                if month > 12:
                    if month == 13:
                        current_year += 1
                    month -= 12

                days_in_month = monthrange(current_year, month)[1]
                days += days_in_month

            interval_type = "days"
            interval = days
        interval_in_seconds = _secondsConverter[interval_type](interval)
        return interval_in_seconds
