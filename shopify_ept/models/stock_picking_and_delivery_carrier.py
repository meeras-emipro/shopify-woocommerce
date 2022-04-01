from odoo import models, fields, api, _


class DeliveryCarrier(models.Model):
    _inherit = "delivery.carrier"

    # This field use to identify the Shopify delivery method
    shopify_code = fields.Char("Shopify Delivery Code")
    shopify_source = fields.Char("Shopify Delivery Source",help="when system will import order then search carrier based on shopift_tracking_company if not found thne search based on source.")
    shopify_tracking_company = fields.Selection([
        ('4PX', '4PX'),
        ('APC', 'APC'),
        ('Amazon Logistics UK', 'Amazon Logistics UK'),
        ('Amazon Logistics US', 'Amazon Logistics US'),
        ('Anjun Logistics', 'Anjun Logistics'),
        ('Australia Post', 'Australia Post'),
        ('Bluedart', 'Bluedart'),
        ('Canada Post', 'Canada Post'),
        ('Canpar', 'Canpar'),
        ('China Post', 'China Post'),
        ('Chukou1', 'Chukou1'),
        ('Correios', 'Correios'),
        ('Couriers Please', 'Couriers Please'),
        ('DHL Express', 'DHL Express'),
        ('DHL eCommerce', 'DHL eCommerce'),
        ('DHL eCommerce Asia', 'DHL eCommerce Asia'),
        ('DPD', 'DPD'),
        ('DPD Local', 'DPD Local'),
        ('DPD UK', 'DPD UK'),
        ('Delhivery', 'Delhivery'),
        ('Eagle', 'Eagle'),
        ('FSC', 'FSC'),
        ('Fastway Australia', 'Fastway Australia'),
        ('FedEx', 'FedEx'),
        ('GLS', 'GLS'),
        ('GLS (US)', 'GLS (US)'),
        ('Globegistics', 'Globegistics'),
        ('Japan Post (EN)', 'Japan Post (EN)'),
        ('Japan Post (JA)', 'Japan Post (JA)'),
        ('La Poste', 'La Poste'),
        ('New Zealand Post', 'New Zealand Post'),
        ('Newgistics', 'Newgistics'),
        ('PostNL', 'PostNL'),
        ('PostNord', 'PostNord'),
        ('Purolator', 'Purolator'),
        ('Royal Mail', 'Royal Mail'),
        ('SF Express', 'SF Express'),
        ('SFC Fulfillment', 'SFC Fulfillment'),
        ('Sagawa (EN)', 'Sagawa (EN)'),
        ('Sagawa (JA)', 'Sagawa (JA)'),
        ('Sendle', 'Sendle'),
        ('Singapore Post', 'Singapore Post'),
        ('StarTrack', 'StarTrack'),
        ('TNT', 'TNT'),
        ('Toll IPEC', 'Toll IPEC'),
        ('UPS', 'UPS'),
        ('USPS', 'USPS'),
        ('Whistl', 'Whistl'),
        ('Yamato (EN)', 'Yamato (EN)'),
        ('Yamato (JA)', 'Yamato (JA)'),
        ('YunExpress', 'YunExpress')
    ], help="shopify_tracking_company selection help:When creating a fulfillment for a supported carrier, send the tracking_company exactly as written in the list above. If the tracking company doesn't match one of the supported entries, then the shipping status might not be updated properly during the fulfillment process.")

    # This method is old flow of delivery method process
    # def shopify_search_create_delivery_carrier(self, line):
    #     delivery_method = line.get('source')
    #     delivery_title = line.get('title')
    #     carrier = False
    #     if delivery_method:
    #         carrier = self.search(
    #                 [('shopify_code', '=', delivery_method)], limit=1)
    #         if not carrier:
    #             carrier = self.search(
    #                     ['|', ('name', '=', delivery_title),
    #                      ('shopify_code', '=', delivery_method)], limit=1)
    #         if not carrier:
    #             carrier = self.search(
    #                     ['|', ('name', 'ilike', delivery_title),
    #                      ('shopify_code', 'ilike', delivery_method)], limit=1)
    #         if not carrier:
    #             product_template = self.env['product.template'].search(
    #                     [('name', '=', delivery_title), ('type', '=', 'service')], limit=1)
    #             if not product_template:
    #                 product_template = self.env['product.template'].create(
    #                         {'name':delivery_title, 'type':'service'})
    #             carrier = self.create(
    #                     {'name':delivery_title, 'shopify_code':delivery_method,
    #                      'product_id':product_template.product_variant_ids[0].id})
    #     return carrier

    def shopify_search_create_delivery_carrier(self,line,instance):
        delivery_source = line.get('source')
        delivery_code = line.get('code')
        delivery_title = line.get('title')
        carrier = self.env['delivery.carrier']
        if delivery_source and delivery_code:
            carrier = self.search(
                [('shopify_source', '=', delivery_source),'|',('shopify_code', '=', delivery_code),('shopify_tracking_company', '=', delivery_code)], limit=1)
            if not carrier:
                carrier = self.search(
                    [('name', '=', delivery_title)], limit=1)
                if carrier:
                    carrier.write({'shopify_source':delivery_source,'shopify_code':delivery_code})
            if not carrier:
                #Added by Harsh Parekh on 11/10/2020 when instance has a shipping product then it will create carrier according that. Task Id:166790
                if not instance.shopify_shipping_product_id:
                    product_template = self.env['product.template'].search(
                        [('name', '=', delivery_title), ('type', '=', 'service')], limit=1)
                    if not product_template:
                        product_template = self.env['product.template'].create(
                            {'name': delivery_title, 'type': 'service'})
                carrier = self.create(
                    {'name': delivery_title, 'shopify_code': delivery_code, 'shopify_source': delivery_source,
                     'product_id':instance.shopify_shipping_product_id.id if instance.shopify_shipping_product_id  else product_template.product_variant_ids[0].id})
        return carrier