import base64
import csv
import logging
import io
import os
import xlrd
from csv import DictWriter
from datetime import datetime
from io import StringIO, BytesIO
from odoo.tools.misc import xlsxwriter
from odoo import models, fields, _
from odoo.exceptions import Warning, ValidationError, UserError

_logger = logging.getLogger("Shopify")


class PrepareProductForExport(models.TransientModel):
    """
    Model for adding Odoo products into Shopify Layer.
    @author: Maulik Barad on Date 11-Apr-2020.
    """
    _name = "shopify.prepare.product.for.export.ept"
    _description = "Prepare product for export in Shopify"

    export_method = fields.Selection([("direct", "Export in Odoo Shopify Product List"),
                                      ("csv", "Export in CSV file"),
                                      ("xlsx", "Export in XLSX file")], default="direct")
    shopify_instance_id = fields.Many2one("shopify.instance.ept")
    datas = fields.Binary("File")
    choose_file = fields.Binary(filters="*.csv", help="Select CSV file to upload.")
    file_name = fields.Char(string="File Name", help="Name of CSV file.")

    def prepare_product_for_export(self):
        """
        This method is used to export products in Shopify layer as per selection.
        If "direct" is selected, then it will direct export product into Shopify layer.
        If "csv" is selected, then it will export product data in CSV file, if user want to do some
        modification in name, description, etc. before importing into Shopify.
        """
        _logger.info("Starting product exporting via %s method..." % self.export_method)

        active_template_ids = self._context.get("active_ids", [])
        templates = self.env["product.template"].browse(active_template_ids)
        product_templates = templates.filtered(lambda template: template.type == "product")
        if not product_templates:
            raise Warning("It seems like selected products are not Storable products.")

        if self.export_method == "direct":
            return self.export_direct_in_shopify(product_templates)
        elif self.export_method == "csv":
            return self.export_csv_file(product_templates)
        else:
            return self.export_xlsx_file(product_templates)

    def export_direct_in_shopify(self, product_templates):
        """
        Creates new product or updates existing product in Shopify layer.
        @author: Maulik Barad on Date 11-Apr-2020.
        Changes done by Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
        Task_id: 181893 - Shopify fixes as per the new version
        """
        shopify_template_id = False
        sequence = 0
        variants = product_templates.product_variant_ids
        shopify_instance = self.shopify_instance_id

        for variant in variants:
            if not variant.default_code:
                continue
            product_template = variant.product_tmpl_id
            if product_template.attribute_line_ids and len(product_template.attribute_line_ids.filtered(
                    lambda x: x.attribute_id.create_variant == "always")) > 3:
                continue
            shopify_template, sequence, shopify_template_id = self.create_or_update_shopify_layer_template(
                shopify_instance, product_template, variant, shopify_template_id, sequence)

            self.create_shopify_template_images(shopify_template)

            if shopify_template and shopify_template.shopify_product_ids and \
                    shopify_template.shopify_product_ids[0].sequence:
                sequence += 1

            shopify_variant = self.create_or_update_shopify_layer_variant(variant, shopify_template_id,
                                                                          shopify_instance, shopify_template, sequence)

            self.create_shopify_variant_images(shopify_template, shopify_variant)
        return True

    def create_or_update_shopify_layer_template(self, shopify_instance, product_template, variant,
                                                shopify_template_id, sequence):
        """ This method is used to create or update the Shopify layer template.
            @return: shopify_template, sequence, shopify_template_id
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
            Task_id: 181893 - Shopify fixes as per the new version
        """
        shopify_templates = shopify_template_obj = self.env["shopify.product.template.ept"]

        shopify_template = shopify_template_obj.search([
            ("shopify_instance_id", "=", shopify_instance.id),
            ("product_tmpl_id", "=", product_template.id)], limit=1)

        if not shopify_template:
            shopify_product_template_vals = self.prepare_template_val_for_export_product_in_layer(product_template,
                                                                                                  shopify_instance,
                                                                                                  variant)
            shopify_template = shopify_template_obj.create(shopify_product_template_vals)
            sequence = 1
            shopify_template_id = shopify_template.id
        else:
            if shopify_template_id != shopify_template.id:
                shopify_product_template_vals = self.prepare_template_val_for_export_product_in_layer(product_template,
                                                                                                      shopify_instance,
                                                                                                      variant)
                shopify_template.write(shopify_product_template_vals)
                shopify_template_id = shopify_template.id
        if shopify_template not in shopify_templates:
            shopify_templates += shopify_template

        return shopify_template, sequence, shopify_template_id

    def prepare_template_val_for_export_product_in_layer(self, product_template, shopify_instance, variant):
        """ This method is used to prepare a template Vals for export/update product
            from Odoo products to the Shopify products layer.
            :param product_template: Record of odoo template.
            :param product_template: Record of instance.
            @return: template_vals
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
            Task_id: 181893 - Shopify fixes as per the new version
        """
        ir_config_parameter_obj = self.env["ir.config_parameter"]
        template_vals = {"product_tmpl_id": product_template.id,
                         "shopify_instance_id": shopify_instance.id,
                         "shopify_product_category": product_template.categ_id.id,
                         "name": product_template.name}
        if ir_config_parameter_obj.sudo().get_param("shopify_ept.set_sales_description"):
            template_vals.update({"description": variant.description_sale})
        return template_vals

    def prepare_variant_val_for_export_product_in_layer(self, shopify_instance, shopify_template, variant, sequence):
        """ This method is used to prepare a vals for the variants.
            @return: shopify_variant_vals
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
            Task_id: 181893 - Shopify fixes as per the new version
        """
        shopify_variant_vals = ({
            "shopify_instance_id": shopify_instance.id,
            "product_id": variant.id,
            "shopify_template_id": shopify_template.id,
            "default_code": variant.default_code,
            "name": variant.name,
            "sequence": sequence
        })
        return shopify_variant_vals

    def create_or_update_shopify_layer_variant(self, variant, shopify_template_id, shopify_instance,
                                               shopify_template, sequence):
        """ This method is used to create/update the variant in the shopify layer.
            @return: shopify_variant
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
            Task_id: 181893 - Shopify fixes as per the new version
        """
        shopify_product_obj = self.env["shopify.product.product.ept"]

        shopify_variant = shopify_product_obj.search([
            ("shopify_instance_id", "=", self.shopify_instance_id.id),
            ("product_id", "=", variant.id),
            ("shopify_template_id", "=", shopify_template_id)])

        shopify_variant_vals = self.prepare_variant_val_for_export_product_in_layer(shopify_instance,
                                                                                    shopify_template, variant,
                                                                                    sequence)
        if not shopify_variant:
            shopify_variant = shopify_product_obj.create(shopify_variant_vals)
        else:
            shopify_variant.write(shopify_variant_vals)

        return shopify_variant

    def prepare_product_data_for_file(self, product_templates):
        """
        This method is use to prepare product data for export csv/xlsx file.
        @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021 .
        Task_id: 181893 - Shopify fixes as per the new version
        """
        product_data_list = []
        for template in product_templates:
            if template.attribute_line_ids and len(
                    template.attribute_line_ids.filtered(lambda x: x.attribute_id.create_variant == "always")) > 3:
                continue
            if len(template.product_variant_ids.ids) == 1 and not template.default_code:
                continue
            for product in template.product_variant_ids.filtered(lambda variant: variant.default_code):
                product_data = self.prepare_row_data_for_file(template, product)
                product_data_list.append(product_data)

        if not product_data_list:
            raise UserError(_("No data found to be exported.\n\nPossible Reasons:\n   - Number of "
                              "attributes are more than 3.\n   - SKU(s) are not set properly."))
        return product_data_list

    def export_csv_file(self, product_templates):
        """
        This method is used for export the odoo products in csv file format
        @param self: It contain the current class Instance
        @author: Nilesh Parmar @Emipro Technologies Pvt. Ltd on date 04/11/2019
        Changes done by Meera Sidapara on date 31/12/2021.
        """
        product_data = self.prepare_product_data_for_file(product_templates)
        buffer = StringIO()
        delimiter = ","
        field_names = list(product_data[0].keys())
        csv_writer = DictWriter(buffer, field_names, delimiter=delimiter)
        csv_writer.writer.writerow(field_names)
        csv_writer.writerows(product_data)
        buffer.seek(0)
        file_data = buffer.read().encode()
        self.write({
            "choose_file": base64.encodebytes(file_data),
            "file_name": "Shopify_export_product_"
        })

        return {
            "type": "ir.actions.act_url",
            "url": "web/content/?model=shopify.prepare.product.for.export.ept&id=%s&field=choose_file&download=true&"
                   "filename=%s.csv" % (self.id, self.file_name + str(datetime.now().strftime("%d/%m/%Y:%H:%M:%S"))),
            "target": self
        }

    def export_xlsx_file(self, product_templates):
        """
        This method is use to export the product data in xlsx file.
        @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021 .
        Task_id: 181893 - Shopify fixes as per the new version
        """
        product_data = self.prepare_product_data_for_file(product_templates)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Map Product')
        header = list(product_data[0].keys())
        header_format = workbook.add_format({'bold': True, 'font_size': 10})
        general_format = workbook.add_format({'font_size': 10})
        worksheet.write_row(0, 0, header, header_format)
        index = 0
        for product in product_data:
            index += 1
            worksheet.write_row(index, 0, list(product.values()), general_format)
        workbook.close()
        b_data = base64.b64encode(output.getvalue())
        self.write({
            "choose_file": b_data,
            "file_name": "Shopify_export_product_"
        })
        return {
            "type": "ir.actions.act_url",
            "url": "web/content/?model=shopify.prepare.product.for.export.ept&id=%s&field=choose_file&download=true&"
                   "filename=%s.xlsx" % (self.id, self.file_name + str(datetime.now().strftime("%d/%m/%Y:%H:%M:%S"))),
            "target": self
        }

    def prepare_row_data_for_file(self, template, product):
        """ This method is used to prepare a row data of csv/xlsx file.
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
            Task_id: 181893 - Shopify fixes as per the new version
        """
        row = {
            "template_name": template.name,
            "product_name": product.name,
            "product_default_code": product.default_code,
            "shopify_product_default_code": product.default_code,
            "product_description": product.description_sale or None,
            "PRODUCT_TEMPLATE_ID": template.id,
            "PRODUCT_ID": product.id,
            "CATEGORY_ID": template.categ_id.id
        }
        return row

    def create_shopify_template_images(self, shopify_template):
        """
        For adding all odoo images into shopify layer only for template.
        @author: Meera Sidapara on Date 31-Dec-2021.
        """
        shopify_product_image_list = []
        shopify_product_image_obj = self.env["shopify.product.image.ept"]

        product_template = shopify_template.product_tmpl_id
        for odoo_image in product_template.ept_image_ids.filtered(lambda x: not x.product_id):
            shopify_product_image = shopify_product_image_obj.search_read(
                [("shopify_template_id", "=", shopify_template.id),
                 ("odoo_image_id", "=", odoo_image.id)], ["id"])
            if not shopify_product_image:
                shopify_product_image_list.append({
                    "odoo_image_id": odoo_image.id,
                    "shopify_template_id": shopify_template.id
                })
        if shopify_product_image_list:
            shopify_product_image_obj.create(shopify_product_image_list)
        return True

    def create_shopify_variant_images(self, shopify_template, shopify_variant):
        """
        For adding first odoo image into shopify layer for variant.
        @author: Meera Sidapara on Date 31-Dec-2021.
        """
        shopify_product_image_obj = self.env["shopify.product.image.ept"]
        product_id = shopify_variant.product_id
        odoo_image = product_id.ept_image_ids
        if odoo_image:
            shopify_product_image = shopify_product_image_obj.search_read(
                [("shopify_template_id", "=", shopify_template.id),
                 ("shopify_variant_id", "=", shopify_variant.id),
                 ("odoo_image_id", "=", odoo_image[0].id)], ["id"])
            if not shopify_product_image:
                shopify_product_image_obj.create({
                    "odoo_image_id": odoo_image[0].id,
                    "shopify_variant_id": shopify_variant.id,
                    "shopify_template_id": shopify_template.id,
                    "sequence": 0
                })
        return True

    def import_products_from_file(self):
        """
        This method is use to import product from csv,xlsx,xls.
        @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
        Task_id: 181893 - Shopify fixes as per the new version
        """
        try:
            if os.path.splitext(self.file_name)[1].lower() not in ['.csv', '.xls', '.xlsx']:
                raise UserError(_("Invalid file format. You are only allowed to upload .csv, .xlsx file."))
            if os.path.splitext(self.file_name)[1].lower() == '.csv':
                self.import_products_from_csv()
            else:
                self.import_products_from_xlsx()
        except Exception as error:
            raise UserError(_("Receive the error while import file. %s") % error)

    def import_products_from_csv(self):
        """
        This method used to import product using csv file in shopify third layer
        images related changes taken by Maulik Barad
        @param : self
        @author: Nilesh Parmar @Emipro Technologies Pvt. Ltd on date 05/11/2019
        Changes done by Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
        Task_id: 181893 - Shopify fixes as per the new version
        """
        file_data = self.read_file()
        self.validate_required_csv_header(file_data.fieldnames)
        self.create_products_from_file(file_data)
        return True

    def import_products_from_xlsx(self):
        """
        This method used to import product using xlsx file in shopify layer.
        @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021 .
        Task_id: 181893 - Shopify fixes as per the new version
        """
        header, product_data = self.read_xlsx_file()
        self.validate_required_csv_header(header)
        self.create_products_from_file(product_data)
        return True

    def validate_required_csv_header(self, header):
        """ This method is used to validate required csv header while csv file import for products.
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021.
            Task_id: 181893 - Shopify fixes as per the new version
        """
        required_fields = ["template_name", "product_name", "product_default_code",
                           "shopify_product_default_code", "product_description",
                           "PRODUCT_TEMPLATE_ID", "PRODUCT_ID", "CATEGORY_ID"]

        for required_field in required_fields:
            if required_field not in header:
                raise UserError(_("Required column is not available in File."))

    def create_products_from_file(self, file_data):
        """
        This method is used to create products in Shopify product layer from the file.
        @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021 .
        Task_id: 181893 - Shopify fixes as per the new version
        """
        prepare_product_for_export_obj = self.env["shopify.prepare.product.for.export.ept"]
        common_log_obj = self.env["common.log.book.ept"]
        common_log_line_obj = self.env["common.log.lines.ept"]
        model_id = common_log_line_obj.get_model_id("shopify.product.product.ept")
        instance = self.shopify_instance_id
        log_book_id = common_log_obj.create({"type": "import",
                                             "module": "shopify_ept",
                                             "shopify_instance_id": instance.id if instance else False,
                                             "model_id": model_id,
                                             "active": True})
        sequence = 0
        row_no = 0
        shopify_template_id = False
        for record in file_data:
            row_no += 1
            message = ""
            if not record["PRODUCT_TEMPLATE_ID"] or not record["PRODUCT_ID"] or not record["CATEGORY_ID"]:
                message += "PRODUCT_TEMPLATE_ID Or PRODUCT_ID Or CATEGORY_ID Not As Per Odoo Product in file at row " \
                           "%s " % row_no
                vals = {"message": message,
                        "model_id": model_id,
                        "log_book_id": log_book_id.id}
                common_log_line_obj.create(vals)
                continue

            shopify_template, shopify_template_id, sequence = self.create_or_update_shopify_template_from_csv(instance,
                                                                                                              record,
                                                                                                              shopify_template_id,
                                                                                                              sequence)

            shopify_variant = self.create_or_update_shopify_variant_from_csv(instance, record, shopify_template_id,
                                                                             sequence)
            prepare_product_for_export_obj.create_shopify_variant_images(shopify_template, shopify_variant)

        if not log_book_id.log_lines:
            log_book_id.unlink()
        return True

    def create_or_update_shopify_template_from_csv(self, instance, record, shopify_template_id, sequence):
        """ This method is used to create or update shopify template while process from csv file import.
            :param record: One row data of csv file.
            @return: shopify_template, shopify_template_id, sequence
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021 .
            Task_id: 181893 - Shopify fixes as per the new version
        """
        shopify_product_template = self.env["shopify.product.template.ept"]
        prepare_product_for_export_obj = self.env["shopify.prepare.product.for.export.ept"]
        shopify_template = shopify_product_template.search(
            [("shopify_instance_id", "=", instance.id),
             ("product_tmpl_id", "=", int(record["PRODUCT_TEMPLATE_ID"]))], limit=1)

        shopify_product_template_vals = {"product_tmpl_id": int(record["PRODUCT_TEMPLATE_ID"]),
                                         "shopify_instance_id": instance.id,
                                         "shopify_product_category": int(record["CATEGORY_ID"]),
                                         "name": record["template_name"]}
        if self.env["ir.config_parameter"].sudo().get_param("shopify_ept.set_sales_description"):
            shopify_product_template_vals.update({"description": record["product_description"]})
        if not shopify_template:
            shopify_template = shopify_product_template.create(shopify_product_template_vals)
            sequence = 1
            shopify_template_id = shopify_template.id
        elif shopify_template_id != shopify_template.id:
            shopify_template.write(shopify_product_template_vals)
            shopify_template_id = shopify_template.id

        prepare_product_for_export_obj.create_shopify_template_images(shopify_template)

        if shopify_template and shopify_template.shopify_product_ids and \
                shopify_template.shopify_product_ids[0].sequence:
            sequence += 1

        return shopify_template, shopify_template_id, sequence

    def create_or_update_shopify_variant_from_csv(self, instance, record, shopify_template_id, sequence):
        """ This method is used to create or update Shopify variants while processing from CSV file import operation.
            @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021 .
            Task_id: 181893 - Shopify fixes as per the new version
        """
        shopify_product_obj = self.env["shopify.product.product.ept"]
        shopify_variant = shopify_product_obj.search(
            [("shopify_instance_id", "=", instance.id),
             ("product_id", "=", int(record["PRODUCT_ID"])),
             ("shopify_template_id", "=", shopify_template_id)])
        shopify_variant_vals = {"shopify_instance_id": instance.id,
                                "product_id": int(record["PRODUCT_ID"]),
                                "shopify_template_id": shopify_template_id,
                                "default_code": record["shopify_product_default_code"],
                                "name": record["product_name"],
                                "sequence": sequence}
        if not shopify_variant:
            shopify_variant = shopify_product_obj.create(shopify_variant_vals)
        else:
            shopify_variant.write(shopify_variant_vals)

        return shopify_variant

    def read_file(self):
        """
        This method reads .csv file
        @author: Nilesh Parmar @Emipro Technologies Pvt. Ltd on date 08/11/2019
        """
        self.write({"datas": self.choose_file})
        self._cr.commit()
        import_file = BytesIO(base64.decodestring(self.datas))
        file_read = StringIO(import_file.read().decode())
        reader = csv.DictReader(file_read, delimiter=",")
        return reader

    def read_xlsx_file(self):
        """
        This method is use to read the xlsx file data.
        @author: Meera Sidapara @Emipro Technologies Pvt. Ltd on date 31 December 2021 .
        Task_id: 181893 - Shopify fixes as per the new version
        """
        validation_header = []
        product_data = []
        sheets = xlrd.open_workbook(file_contents=base64.b64decode(self.choose_file.decode('UTF-8')))
        header = dict()
        is_header = False
        for sheet in sheets.sheets():
            for row_no in range(sheet.nrows):
                if not is_header:
                    headers = [d.value for d in sheet.row(row_no)]
                    validation_header = headers
                    [header.update({d: headers.index(d)}) for d in headers]
                    is_header = True
                    continue
                row = dict()
                [row.update({k: sheet.row(row_no)[v].value}) for k, v in header.items() for c in
                 sheet.row(row_no)]
                product_data.append(row)
        return validation_header, product_data
