from odoo import models, fields, api, _
from datetime import datetime, timedelta
import json
import logging
import time

_logger = logging.getLogger("shopify_customer_queue_line_process")


class ShopifyCustomerDataQueueLineEpt(models.Model):
    _name = "shopify.customer.data.queue.line.ept"
    _description = 'Shopify Synced Customer Data Line'

    state = fields.Selection([('draft', 'Draft'), ('failed', 'Failed'), ('done', 'Done'),
                              ("cancel", "Cancelled")],
                             default='draft')
    shopify_synced_customer_data = fields.Char(string='Shopify Synced Data')
    shopify_customer_data_id = fields.Text(string='Customer ID')
    synced_customer_queue_id = fields.Many2one("shopify.customer.data.queue.ept", string="Shopify Customer",
                                          ondelete="cascade")
    last_process_date = fields.Datetime('Last Process Date', readonly=True)
    shopify_instance_id = fields.Many2one('shopify.instance.ept', string='Instance')
    common_log_lines_ids = fields.One2many("common.log.lines.ept",
                                           "shopify_customer_data_queue_line_id",
                                           help="Log lines created against which line.")

    name = fields.Char(string="Customer", help="Shopify Customer Name")

    def to_process_customer_child_cron(self):
        """This method used to start the child process cron for process the synced shopify customer data.
            @param : self
            @author: Angel Patel @Emipro Technologies Pvt.Ltd on date 25/10/2019.
            :Task ID: 157065
        """
        self.sync_shopify_customer_into_odoo()
        # child_cron_of_process = self.env.ref('shopify_ept.ir_cron_child_to_process_shopify_synced_customer_data')
        # if child_cron_of_process and not child_cron_of_process.active:
        #     results = self.search([('state', '=', 'draft')], limit=100)
        #     if not results:
        #         return True
        #     child_cron_of_process.write({'active': True,
        #                                  'numbercall': 1,
        #                                  'nextcall': datetime.now() + timedelta(seconds=10)
        #                                  })
        return True

    @api.model
    def sync_shopify_customer_into_odoo(self, results=False):
        """
        Change the queue and queue line record state using this compute method
        :param results:
        :return:
        :author: Angel Patel @Emipro Technologies Pvt.Ltd on date 02/11/2019.
        :Task ID: 157065
        """
        partner_obj = self.env['res.partner']
        common_log_obj = self.env["common.log.book.ept"]
        customer_data_queue_obj = self.env['shopify.customer.data.queue.ept']
        customer_queue_ids = []
        if not results:
            query = """select queue.id
                from shopify_customer_data_queue_line_ept as queue_line
                inner join shopify_customer_data_queue_ept as queue on queue_line.synced_customer_queue_id = queue.id
                where queue_line.state='draft' and queue.is_action_require = 'False'
                ORDER BY queue_line.create_date ASC"""
            self._cr.execute(query)
            customer_data_queue_list = self._cr.fetchall()
            if not customer_data_queue_list:
                return
            for customer_data_queue_id in customer_data_queue_list:
                if customer_data_queue_id[0] not in customer_queue_ids:
                    customer_queue_ids.append(customer_data_queue_id[0])
            queue = customer_data_queue_obj.browse(customer_queue_ids[0])
            start = time.time()
            customer_queue_process_cron_time = queue.shopify_instance_id.get_shopify_cron_execution_time(
            "shopify_ept.ir_cron_parent_to_process_shopify_synced_customer_data")
            for queue_id in customer_queue_ids:
                queue = customer_data_queue_obj.browse(queue_id)
                results = queue.synced_customer_queue_line_ids
                # For counting the queue crashes and creating schedule activity for the queue.
                queue.queue_process_count += 1
                if queue.queue_process_count > 3:
                    queue.is_action_require = True
                    note = "<p>Attention %s queue is processed 3 times you need to process it manually</p>" % (queue.name)
                    queue.message_post(body=note)
                    if queue.shopify_instance_id.is_shopify_create_schedule:
                        model_id = self.env['ir.model'].search([('model', '=', 'shopify.customer.data.queue.ept')]).id
                        common_log_obj.create_crash_queue_schedule_activity(queue, model_id, note)
                    return
                if time.time() - start > customer_queue_process_cron_time - 60:
                    return True
                # Below three line add by haresh mori on data 21/1/2020. To bypass the process when
                # the instance is not active.
                if not queue.shopify_instance_id.active:
                    _logger.info("Instance '{}' is not active.".format(queue.shopify_instance_id.name))
                    return True
                if queue.common_log_book_id:
                    log_book_id = queue.common_log_book_id
                else:
                    log_book_id = common_log_obj.create({'type': 'import',
                                                         'module': 'shopify_ept',
                                                         'shopify_instance_id': queue.shopify_instance_id.id,
                                                         'active': True})
                    # below two line add by Dipak Gogiya on date 15/01/2020, this is used to update
                    # is_process_queue as False.
                    self.env.cr.execute(
                        """update shopify_product_data_queue_ept set is_process_queue = False where is_process_queue = True""")
                    self._cr.commit()
                commit_count = 0
                for line in results:
                    commit_count += 1
                    # Added by Dipak gogiya
                    if commit_count == 10:
                        if queue:
                            queue.is_process_queue = True
                        self._cr.commit()
                        commit_count = 0
                    instance = line.synced_customer_queue_id.shopify_instance_id
                    partner = partner_obj.create_or_update_customer(
                        vals=json.loads(line.shopify_synced_customer_data), is_company=True, parent_id=False, type=False,
                        instance=instance, customer_data_queue_line_id=line, log_book_id=log_book_id)
                    if partner:
                        line.update({'state': 'done'})
                    else:
                        line.update({'state': 'failed'})
                    # Below two-line add by Dipak Gogiya on date 15/01/2020 to manage the which queue is running in the background
                    if queue:
                        queue.is_process_queue = False
                # _logger.info("Commit 100 shopify customer queue line")
                # self._cr.commit()
                queue.common_log_book_id = log_book_id
                # draft_or_failed_queue_line = results.filtered(lambda line: line.state in ['draft', 'failed'])
                # if draft_or_failed_queue_line:
                #     queue_id.write({'state': "partially_completed"})
                # else:
                #     queue_id.write({'state': "completed"})
                if queue.common_log_book_id and not queue.common_log_book_id.log_lines:
                    queue.common_log_book_id.unlink()

    def create_customer_queue_schedule_activity(self, queue_id):
        """
        this method is used to create a schedule activity for queue.
        @:parameter : queue_id : it is object of queue
        Author: Nilesh Parmar
        Date: 10 february 2020.
        :return:
        """
        mail_activity_obj = self.env['mail.activity']
        ir_model_obj = self.env['ir.model']
        model_id = ir_model_obj.search([('model', '=', 'shopify.customer.data.queue.ept')])
        activity_type_id = queue_id and queue_id.shopify_instance_id.shopify_activity_type_id.id
        date_deadline = datetime.strftime(
            datetime.now() + timedelta(
                days=int(queue_id.shopify_instance_id.shopify_date_deadline)),
            "%Y-%m-%d")
        if queue_id:
            note = "Attention %s queue is processed 3 times you need to process it manually" % (queue_id.name)
            for user_id in queue_id.shopify_instance_id.shopify_user_ids:
                mail_activity = mail_activity_obj.search(
                    [('res_model_id', '=', model_id.id), ('user_id', '=', user_id.id),
                     ('res_name', '=', queue_id.name),
                     ('activity_type_id', '=', activity_type_id)])
                note_2 = '<p>' + note + '</p>'
                if not mail_activity or mail_activity.note != note_2:
                    vals = {'activity_type_id': activity_type_id,
                            'note': note,
                            'res_id': queue_id.id,
                            'user_id': user_id.id or self._uid,
                            'res_model_id': model_id.id,
                            'date_deadline': date_deadline}
                    try:
                        mail_activity_obj.create(vals)
                    except:
                        _logger.info(
                            "Unable to create schedule activity, Please give proper "
                            "access right of this user :%s  " % user_id.name)
                        pass
        return True
